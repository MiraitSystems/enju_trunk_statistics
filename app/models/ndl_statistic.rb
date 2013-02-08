# -*- encoding: utf-8 -*-
class NdlStatistic < ActiveRecord::Base
  has_many :ndl_stat_manifestations, :dependent => :destroy
  has_many :ndl_stat_accepts, :dependent => :destroy
  has_many :ndl_stat_checkouts, :dependent => :destroy
  has_many :ndl_stat_jma_publications, :dependent => :destroy
  attr_accessible :term_id
  
  term_ids = Term.select(:id).map(&:id)
  
  validates_presence_of :term_id
  validates_uniqueness_of :term_id
  validates_inclusion_of :term_id, :in => term_ids

  # NDL 年報用集計処理
  def calc_all
    # validates term_id
    begin
      @prev_term_end = Term.where(:id => term_id).first.start_at.yesterday
      @curr_term_end = Term.where(:id => term_id).first.end_at
    rescue Exception => e
      p "Failed to get Term object for #{term_id}: #{e}"
      logger.error "Failed to get Term object for #{term_id}: #{e}"
      return false
    end
    # calculate ndl statistics
    self.calc_manifestation_counts
    self.calc_accept_counts
    self.calc_checkout_counts
    self.aggregate_jma_publications
  rescue Exception => e
    p "Failed to calculate ndl statistics: #{e}"
    logger.error "Failed to calculate ndl statistics: #{e}"
  end
  
  # 1. 所蔵
  def calc_manifestation_counts
    NdlStatistic.transaction do
      # p "ndl_statistics of manifestation_counts"
      # 書籍、逐次刊行物
      [ "book", "magazine" ].each do |type|
        book_type = (type == 'book') ? ['%book', '%monograph'] : ['%magazine', '%serial_book']
        items_all = Item.
          includes(:manifestation => [:manifestation_type, :carrier_type]).
          where("manifestation_types.name like ? OR
                 manifestation_types.name like ?", book_type[0], book_type[1]).
          where("carrier_types.name = 'print'").
          where("bookbinder_id IS NULL OR items.bookbinder IS TRUE")
	# 国内, 国外
        [ "domestic", "foreign" ].each do |region|
	  manifestation_type = (region == "domestic") ? 'japanese%' : 'foreign%'
	  items = items_all.
	    where("manifestation_types.name like ?", manifestation_type)
	  # 前年度末現在数
	  prev = items.includes(:circulation_status).
                   where("circulation_statuses.name not in ('Removed', 'Lost', 'Missing')").
	           where("items.created_at < ?", @prev_term_end).count
	  # 本年度増加数
	  inc = items.includes(:circulation_status).
                   where("circulation_statuses.name != 'Missing'").
                   where("items.created_at between ? and ?",
                         @prev_term_end, @curr_term_end).count
	  # 本年度減少数
	  dec = items.includes(:circulation_status).
                   where("circulation_statuses.name in ('Removed', 'Lost')").
                   where("items.removed_at between ? and ?",
                         @prev_term_end, @curr_term_end).count
	  # 本年度末現在数
	  curr = items.includes(:circulation_status).
                   where("circulation_statuses.name not in ('Removed', 'Lost', 'Missing')").
	           where("items.created_at < ?", @curr_term_end).count
	  # サブクラス生成
          ndl_stat_manifestations.create(
            :item_type => type,
	    :region => region,
	    :previous_term_end_count => prev,
	    :inc_count => inc,
	    :dec_count => dec,
	    :current_term_end_count => curr
          )
	end
      end
      # その他
      [ "other_micro", "other_av", "other_file" ].each do |type|
        case type
	when "other_micro"
	  # マイクロ資料
          items = Item.
            includes(:manifestation => :carrier_type).
            where("carrier_types.name = 'micro'").
            where("bookbinder_id IS NULL OR items.bookbinder IS TRUE")
	when "other_av"
	  # 視聴覚資料
          items = Item.
            includes(:manifestation => :carrier_type).
            where("carrier_types.name in ('CD','DVD','AV')").
            where("bookbinder_id IS NULL OR items.bookbinder IS TRUE")
	when "other_file"
	  # 電子出版物
          items = Item.
            includes(:manifestation => :carrier_type).
            where("carrier_types.name = 'file'").
            where("bookbinder_id IS NULL OR items.bookbinder IS TRUE")
	end
        region = "na"
        # 前年度末現在数
	prev = items.includes(:circulation_status).
                 where("circulation_statuses.name not in ('Removed', 'Lost', 'Missing')").
	         where("items.created_at < ?", @prev_term_end).count
	# 本年度増加数
	inc = items.includes(:circulation_status).
                 where("circulation_statuses.name != 'Missing'").
                 where("items.created_at between ? and ?",
                       @prev_term_end, @curr_term_end).count
	# 本年度減少数
	dec = items.includes(:circulation_status).
                 where("circulation_statuses.name in ('Removed', 'Lost')").
                 where("items.removed_at between ? and ?",
                       @prev_term_end, @curr_term_end).count
	# 本年度末現在数
	curr = items.includes(:circulation_status).
                 where("circulation_statuses.name not in ('Removed', 'Lost', 'Missing')").
	         where("items.created_at < ?", @curr_term_end).count
	# サブクラス生成
        ndl_stat_manifestations.create(
          :item_type => type,
	  :region => region,
	  :previous_term_end_count => prev,
	  :inc_count => inc,
	  :dec_count => dec,
	  :current_term_end_count => curr
        )
      end
    end
  rescue Exception => e
    p "Failed to manifestation counts: #{e}"
    logger.error "Failed to calculate manifestation counts: #{e}"
  end

  # 2. 受入
  def calc_accept_counts
    NdlStatistic.transaction do
      # p "ndl_statistics of accept_counts"
      # 書籍、逐次刊行物
      [ "book", "magazine" ].each do |type|
        book_type = (type == 'book') ? ['%book', '%monograph'] : ['%magazine', '%serial_book']
        items_all = Item.
          includes(:manifestation => [:manifestation_type, :carrier_type]).
          where("manifestation_types.name like ? OR
                 manifestation_types.name like ?", book_type[0], book_type[1]).
          where("carrier_types.name = 'print'").
          where("bookbinder_id IS NULL OR items.bookbinder IS TRUE")
	# 国内, 国外
        [ "domestic", "foreign" ].each do |region|
	  manifestation_type = (region == "domestic") ? 'japanese%' : 'foreign%'
	  items = items_all.
	    where("manifestation_types.name like ?", manifestation_type)
	  # 購入
	  purchase = items.includes(:accept_type).
                       where("accept_types.name = 'purchase'").
	               count
	  # 寄贈
	  donation = items.includes(:accept_type).
                       where("accept_types.name in ('donation','jma','wmo')").
                       count
	  # 生産
	  if type == 'book'
	    production = items.includes(:accept_type).
                           where("accept_types.name = 'production'").
			   count
	  else
	    production = 0
	  end
	  # サブクラス生成
          ndl_stat_accepts.create(
            :item_type => type,
	    :region => region,
	    :purchase => purchase,
	    :donation => donation,
	    :production => production
          )
	end
      end
      # その他
      [ "other_micro", "other_av", "other_file" ].each do |type|
        case type
	when "other_micro"
	  # マイクロ資料
          items = Item.
            includes(:manifestation => :carrier_type).
            where("carrier_types.name = 'micro'").
            where("bookbinder_id IS NULL OR items.bookbinder IS TRUE")
	when "other_av"
	  # 視聴覚資料
          items = Item.
            includes(:manifestation => :carrier_type).
            where("carrier_types.name in ('CD','DVD','AV')").
            where("bookbinder_id IS NULL OR items.bookbinder IS TRUE")
	when "other_file"
	  # 電子出版物
          items = Item.
            includes(:manifestation => :carrier_type).
            where("carrier_types.name = 'file'").
            where("bookbinder_id IS NULL OR items.bookbinder IS TRUE")
	end
        region = "na"
	# 購入
	purchase = items.includes(:accept_type).
                     where("accept_types.name = 'purchase'").
	             count
	# 寄贈
	donation = items.includes(:accept_type).
                     where("accept_types.name in ('donation','jma','wmo')").
                     count
	# サブクラス生成
        ndl_stat_accepts.create(
          :item_type => type,
	  :region => region,
	  :purchase => purchase,
	  :donation => donation,
	  :production => 0
        )
      end
    end
  rescue Exception => e
    p "Failed to accept counts: #{e}"
    logger.error "Failed to accept manifestation counts: #{e}"
  end

  # 3. 利用
  def calc_checkout_counts
    NdlStatistic.transaction do
      # p "ndl_statistics of checkout_counts"
      # 書籍、逐次刊行物
      [ "book", "magazine" ].each do |type|
        book_type = (type == 'book') ? ['%book', '%monograph'] : ['%magazine', '%serial_book']
        checkouts = Checkout.
          joins(:item => { :manifestation => [:manifestation_type, :carrier_type] }).
          where("manifestation_types.name like ? OR
                 manifestation_types.name like ?", book_type[0], book_type[1]).
          where("carrier_types.name = 'print'")
	# 貸出者数
	user = checkouts.where("checkouts.created_at between ? and ?",
	                        @prev_term_end, @curr_term_end).
	                 count(:user_id, :distinct => true)
	# 貸出資料数
	item = checkouts.where("checkouts.created_at between ? and ?",
	                        @prev_term_end, @curr_term_end).
			 where(:checkout_renewal_count => 0).
	                 count
        ndl_stat_checkouts.create(
          :item_type => type,
	  :user => user,
	  :item => item
        )
      end

      # その他
      type = 'other'
      checkouts = Checkout.
        joins(:item => { :manifestation => [:manifestation_type, :carrier_type] }).
        where("manifestation_types.name not like ?", '%book').
        where("manifestation_types.name not like ?", '%monograph').
        where("manifestation_types.name not like ?", '%magazine').
        where("manifestation_types.name not like ?", '%serial_book').
        where("carrier_types.name = 'print'")
      # 貸出者数
      user = checkouts.where("checkouts.created_at between ? and ?",
                             @prev_term_end, @curr_term_end).
	               count(:user_id, :distinct => true)
      # 貸出資料数
      item = checkouts.where("checkouts.created_at between ? and ?",
	                      @prev_term_end, @curr_term_end).
	                      where(:checkout_renewal_count => 0).
                       count
      ndl_stat_checkouts.create(
        :item_type => type,
        :user => user,
	:item => item
      )

    end
  rescue Exception => e
    p "Failed to checkout counts: #{e}"
    logger.error "Failed to calculate checkout counts: #{e}"
  end

  # 7. 刊行資料
  def aggregate_jma_publications
    NdlStatistic.transaction do
      # p "ndl_statistics of jma_publications"
      items = Item.includes(:manifestation, :accept_type).
                   where("accept_types.name = ?", 'jma').
	           where("manifestations.created_at between ? and ?",
	                 @prev_term_end, @curr_term_end)
      items.each do |i|
        # 資料名
        original_title = i.manifestation.original_title
	# 巻号年月次
	number_string = "#{i.manifestation.volume_number_string}巻#{i.manifestation.issue_number_string}号(#{i.manifestation.serial_number_string})"
        ndl_stat_jma_publications.create(
	  :original_title => original_title,
	  :number_string => number_string
        )
      end
    end
  rescue Exception => e
    p "Failed to create jma_publication list: #{e}"
    logger.error "Failed to create jma_publication list: #{e}"
  end

private
  # excel 出力
  def self.get_ndl_report_excelx(ndl_statistic)
    # initialize
    out_dir = "#{Rails.root}/private/system/ndl_report_excelx" 
    excel_filepath = "#{out_dir}/ndlreport#{Time.now.strftime('%s')}#{rand(10)}.xlsx"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)

    logger.info "get_ndl_report_excelx filepath=#{excel_filepath}"
    
    font_size = 10
    height = font_size * 1.5
    
    require 'axlsx'
    Axlsx::Package.new do |p|
      wb = p.workbook
      wb.styles do |s|
        title_style = s.add_style :font_name => Setting.manifestation_list_print_excelx.fontname,
	                          :alignment => { :vertical => :center },
				  :sz => font_size+2, :b => true
        header_style = s.add_style :font_name => Setting.manifestation_list_print_excelx.fontname,
	                           :alignment => { :vertical => :center },
                                   :border => Axlsx::STYLE_THIN_BORDER,
                                   :sz => font_size, :b => true
        default_style = s.add_style :font_name => Setting.manifestation_list_print_excelx.fontname,
	                            :alignment => { :vertical => :center },
                                    :border => Axlsx::STYLE_THIN_BORDER,
				    :sz => font_size

        # 1.所蔵
        wb.add_worksheet(:name => "1. 所蔵") do |sheet|
	  sheet.add_row ['1. 所蔵'], :style => title_style, :height => height*2
	  
	  # (1) 図書
	  sheet.add_row
	  sheet.add_row ['(1) 図書'], :style => title_style, :height => height*2
	  sheet.add_row
	  header = ['','前年度末現在数','本年度増加数(受入数)','本年度減少数(除籍数)','本年度末現在数']
	  sheet.add_row header, :style => header_style, :height => height
	  sheet.column_info.each do |c|
	    c.width = 25
	  end
	  sheet.column_info[0].width = 15
	  ndl_statistic.ndl_stat_manifestations.where(:item_type => 'book').each do |i|
	    row = []
	    region = i.region == 'domestic' ? '国内' : '外国'
	    row << region
	    row << i.previous_term_end_count
	    row << i.inc_count
	    row << i.dec_count
	    row << i.current_term_end_count
	    sheet.add_row row, :style => default_style, :height => height
	  end
	  row = ['計','','','','']
	  sheet.add_row row, :style => default_style, :height => height
	  
	  # (2) 逐次刊行物
	  sheet.add_row
	  sheet.add_row ['(2) 逐次刊行物'], :style => title_style, :height => height*2
	  sheet.add_row
	  header = ['','','前年度末現在数','本年度増加数(受入数)','本年度減少数(除籍数)','本年度末現在数']
	  sheet.add_row header, :style => header_style, :height => height
	  sheet.column_info.each do |c|
	    c.width = 25
	  end
	  sheet.merge_cells("A12:B12")
	  ndl_statistic.ndl_stat_manifestations.where(:item_type => 'book').each do |i|
	    row = []
	    region = i.region == 'domestic' ? '国内' : '外国'
	    row << region
	    row << '雑誌'
	    row << i.previous_term_end_count
	    row << i.inc_count
	    row << i.dec_count
	    row << i.current_term_end_count
	    sheet.add_row row, :style => default_style, :height => height
	    row = ['','新聞(システム管理外)','','','','']
	    sheet.add_row row, :style => default_style, :height => height
	    row = ['','計(種)','','','','']
	    sheet.add_row row, :style => default_style, :height => height
	  end
	  sheet.merge_cells("A13:A15")
	  sheet.merge_cells("A16:A18")
	  row = ['合計(種)','','','','','']
	  sheet.add_row row, :style => default_style, :height => height
	  sheet.merge_cells("A19:B19")
	  
	  # (3) 非図書資料
	  sheet.add_row
	  sheet.add_row ['(3) 非図書資料'], :style => title_style, :height => height*2
	  sheet.add_row
	  header = ['','前年度末現在数','本年度増加数(受入数)','本年度減少数(除籍数)','本年度末現在数']
	  sheet.add_row header, :style => header_style, :height => height
	  sheet.column_info.each do |c|
	    c.width = 25
	  end
	  ndl_statistic.ndl_stat_manifestations.where("item_type like ?", 'other_%').each do |i|
	    row = []
            case i.item_type
            when "other_micro"
	      item_type = 'マイクロ資料'
	    when "other_av"
	      item_type = '視聴覚資料'
	    when "other_file"
	      item_type = '電子出版物'
	    end
	    row << item_type
	    row << i.previous_term_end_count
	    row << i.inc_count
	    row << i.dec_count
	    row << i.current_term_end_count
	    sheet.add_row row, :style => default_style, :height => height
	  end
	  row = ['その他','','','','']
	  sheet.add_row row, :style => default_style, :height => height
	  row = ['合計','','','','']
	  sheet.add_row row, :style => default_style, :height => height
	end

        # 2.受入
        wb.add_worksheet(:name => "2. 受入") do |sheet|
	  sheet.add_row ['2. 受入'], :style => title_style, :height => height*2
	  
	  # (1) 図書
	  sheet.add_row
	  sheet.add_row ['(1) 図書'], :style => title_style, :height => height*2
	  sheet.add_row
	  header = ['','購入','寄贈','管理換','生産','合計']
	  sheet.add_row header, :style => header_style, :height => height
	  sheet.column_info.each do |c|
	    c.width = 15
	  end
	  ndl_statistic.ndl_stat_accepts.where(:item_type => 'book').each do |i|
	    row = []
	    region = i.region == 'domestic' ? '国内' : '外国'
	    row << region
	    row << i.purchase
	    row << i.donation
	    row << ''
	    row << i.production
	    row << ''
	    sheet.add_row row, :style => default_style, :height => height
	  end
	  row = ['合計(冊)','','','','','']
	  sheet.add_row row, :style => default_style, :height => height
	  
	  # (2) 逐次刊行物 (年度末時点)
	  sheet.add_row
	  sheet.add_row ['(2) 逐次刊行物 (年度末時点)'], :style => title_style, :height => height*2
	  sheet.add_row
	  header = ['','購入','','','寄贈','','','合計']
	  sheet.add_row header, :style => header_style, :height => height
	  header = ['','国内','外国','計','国内','外国','計','']
	  sheet.add_row header, :style => header_style, :height => height
	  sheet.merge_cells("A12:A13")
	  sheet.merge_cells("B12:D12")
	  sheet.merge_cells("E12:G12")
	  sheet.merge_cells("H12:H13")
	  sheet.column_info.each do |c|
	    c.width = 15
	  end
	  stats_domestic = ndl_statistic.ndl_stat_accepts.
	                     where(:item_type => 'magazine').
			     where(:region => 'domestic')
	  stats_foreign = ndl_statistic.ndl_stat_accepts.
	                    where(:item_type => 'magazine').
			    where(:region => 'foreign')
	  row = []
	  row << '雑誌'
	  row << stats_domestic.first.purchase
	  row << stats_foreign.first.purchase
	  row << ''
	  row << stats_domestic.first.donation
	  row << stats_foreign.first.donation
	  row << ''
	  row << ''
	  sheet.add_row row, :style => default_style, :height => height
	  row = ['新聞','','','','','','','']
	  sheet.add_row row, :style => default_style, :height => height
	  row = ['合計(種)','','','','','','','']
	  sheet.add_row row, :style => default_style, :height => height
	  
	  # (3) 非図書資料
	  sheet.add_row
	  sheet.add_row ['(3) 非図書資料'], :style => title_style, :height => height*2
	  sheet.add_row
	  header = ['','購入','寄贈','合計']
	  sheet.add_row header, :style => header_style, :height => height
	  sheet.column_info.each do |c|
	    c.width = 25
	  end
	  ndl_statistic.ndl_stat_accepts.where("item_type like ?", 'other_%').each do |i|
	    row = []
            case i.item_type
            when "other_micro"
	      item_type = 'マイクロ資料'
	    when "other_av"
	      item_type = '視聴覚資料'
	    when "other_file"
	      item_type = '電子出版物'
	    end
	    row << item_type
	    row << i.purchase
	    row << i.donation
	    row << ''
	    sheet.add_row row, :style => default_style, :height => height
	  end
	  row = ['その他','','','']
	  sheet.add_row row, :style => default_style, :height => height
	  row = ['合計','','','']
	  sheet.add_row row, :style => default_style, :height => height
	end

        # 3.利用
        wb.add_worksheet(:name => "3. 利用") do |sheet|
	  sheet.add_row ['3. 利用'], :style => title_style, :height => height*2
	  sheet.add_row
	  header = ['','入館者数','閲覧資料数','貸出者数','貸出資料数']
	  sheet.add_row header, :style => header_style, :height => height
	  sheet.column_info.each do |c|
	    c.width = 30
	  end
	  ndl_statistic.ndl_stat_checkouts.each do |i|
	    case i.item_type
	    when 'book'
	      item_type = '図書'
	    when 'magazine'
	      item_type = '雑誌'
	    else
	      item_type = 'その他'
	      row = ['新聞','','','','']
	      sheet.add_row row, :style => default_style, :height => height
	    end
	    row = []
	    row << item_type
	    row << ''
	    row << ''
	    row << i.user
	    row << i.item
	    sheet.add_row row, :style => default_style, :height => height
	  end
	  row = ['計','','','','']
	  sheet.add_row row, :style => default_style, :height => height
	end

        # 7.刊行資料
        wb.add_worksheet(:name => "7. 刊行資料") do |sheet|
	  sheet.add_row ['7. 刊行資料'], :style => title_style, :height => height*2
	  sheet.add_row
	  sheet.add_row ['資料名', '巻号年月次'], :style => header_style, :height => height
	  sheet.column_info[0].width = 60
	  sheet.column_info[1].width = 40
	  ndl_statistic.ndl_stat_jma_publications.each do |i|
	    row = ["#{i.original_title}", "#{i.number_string}"]
	    sheet.add_row row, :style => default_style, :height => height
	  end
	end

        p.serialize(excel_filepath)
      end
      return excel_filepath
    end
  rescue Exception => e
    p "Failed to create ndl report excelxt: #{e}"
    logger.error "Failed to create ndl report excelx: #{e}"
  end

end
