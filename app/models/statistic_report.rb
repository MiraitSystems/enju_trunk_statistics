class StatisticReport < ActiveRecord::Base
  # 要求された書式で統計を生成する
  # -----
  # output.result_type: 生成結果のタイプ
  # output.data: 生成結果のデータ
  # output.job_name: 後で処理する際のジョブ名(result_typeが:delayedのとき)
  def self.generate_report(target, type, current_user, options = {}, &block)
    get_total = proc do
      case target
      when 'users'          then User.count
      when 'manifestations' then Manifestation.without_master.count
      when 'departments'    then Department.count   
      when 'yearly'
        Statistic.where(
          "data_type IN ('121', '133') 
            AND yyyymm >= #{options[:start_at]}04 
            AND yyyymm <= #{options[:end_at].to_i + 1}03"
        ).count
      when 'departments'
        Statistic.where(
          "data_type IN ('122', '133') 
            AND department_id != 0 
            AND yyyymm >= #{options[:term]}04 
            AND yyyymm <= #{options[:term].to_i + 1}03"
        ).count
      else 0
      end
    end
    # 処理数が指定件数以上のとき delayed job に処理を渡す
    threshold ||= Setting.background_job.threshold.export rescue nil
    if threshold and threshold > 0 and get_total.call > threshold
      job_name = GenerateStatisticReportJob.generate_job_name
      Delayed::Job.enqueue GenerateStatisticReportJob.new(job_name, target, type, current_user, options)
      output = OpenStruct.new
      output.result_type = :delayed
      output.job_name    = job_name
      block.call(output)
      return
    end
    generate_report_internal(target, type, options, &block)
  end

  def self.generate_report_internal(target, type, options = {}, &block)
    output = OpenStruct.new
    output.result_type = :data
    output.filename    = "#{target}_report.#{type}"
    
    method = "create_report_#{type}"
    case type
    when 'pdf' then result = output.__send__("#{output.result_type}=", self.__send__(method, target, options).generate)
    when 'tsv' then result = nil # TODO: TSVの処理を書く
    end
    block.call(output)
  end

  def self.create_report_pdf(target, options = {})
    # pdf処理 重複部分の値をセット
    report = ThinReports::Report.new :layout => get_layout_path("#{target}_report")
    report.events.on :page_create do |e| e.page.item(:page).value(e.page.no) end
    report.events.on :generate do |e| 
      e.pages.each do |page| page.item(:total).value(e.report.page_count) end
    end
    report.start_new_page
    report.page.values(:date => Time.now)
    # pdf処理 独自部分の値をセット
    case target
    when 'yearly'         then set_yearly_report_pdf(report, options) 
    when 'users'          then set_users_report_pdf(report, options)
    when 'departments'    then set_departments_report_pdf(report, options) 
    when 'manifestations' then set_manifestations_report_pdf(report, options)
    end
    return report
  end

  # 資料別利用統計
  # TODO: 図書館ごとの統計は未対応
  def self.set_manifestations_report_pdf(report, options)
    # term
    report.page.values(:term => options[:term])
    # list
    manifestations = Manifestation.without_master.order('jpn_or_foreign ASC, manifestation_type_id ASC, ndc ASC, original_title ASC') 
    manifestations.each_with_index do |manifestation, num|
      report.page.list(:list).add_row do |row|
        row.item('jpn_or_foreign').value(Manifestation::JPN_OR_FOREIGN.invert[manifestation.jpn_or_foreign] || I18n.t('jpn_or_foreign.other'))
        row.item('manifestation_type').value(manifestation.manifestation_type.display_name.localize)
        row.item('ndc').value(manifestation.ndc)
        row.item('title').value(manifestation.original_title)
        checkoutall_cnt = 0
        reserveall_cnt = 0
        1.upto(12) do |month|
          yyyymm = "#{month > 3 ? options[:term] : options[:term].to_i + 1}#{"%02d" % month}"
          start_at = Time.zone.parse("#{yyyymm}01").beginning_of_month
          end_at   = Time.zone.parse("#{yyyymm}31").end_of_month
          checkout_cnt = manifestation.items.inject(0) do |sum, item| 
            sum += Checkout.where("item_id = #{item.id} AND checked_at >= '#{start_at}' AND checked_at <= '#{end_at}'").count
          end
          reserve_cnt = Reserve.where("manifestation_id = #{manifestation.id} AND created_at >= '#{start_at}' AND created_at <= '#{end_at}'").count
          row.item("checkout#{month}").value(checkout_cnt)
          row.item("reserve#{month}").value(reserve_cnt)
          checkoutall_cnt += checkout_cnt
          reserveall_cnt  += reserve_cnt
        end
        row.item("checkoutall").value(checkoutall_cnt)
        row.item("reserveall").value(reserveall_cnt)
        # layout
        if manifestation.jpn_or_foreign == manifestations[num - 1].try(:jpn_or_foreign)
          row.item('jpn_or_foreign').hide
          if manifestation.manifestation_type == manifestations[num - 1].try(:manifestation_type)
            row.item('manifestation_type').hide 
            if manifestation.ndc == manifestations[num - 1].try(:ndc)
              row.item('ndc').hide 
              if manifestation.original_title == manifestations[num - 1].try(:original_title)
                row.item('title').hide
              end
            end
          end
        end
        if manifestation.jpn_or_foreign == manifestations[num + 1].try(:jpn_or_foreign)
          row.item('jpn_or_foreign_line').hide 
          if manifestation.manifestation_type == manifestations[num + 1].try(:manifestation_type)
            row.item('manifestation_type_line').hide
            if manifestation.ndc == manifestations[num + 1].try(:ndc)
              row.item('ndc_line').hide
              if manifestation.original_title == manifestations[num - 1].try(:original_title)
                row.item('title_line').show
              end
            end
          end
        end
      end
    end
 end

  # 利用者別利用統計
  def self.set_users_report_pdf(report, options ={})
    # term
    report.page.values(:term => options[:term])
    # list
    users = User.order('library_id ASC, required_role_id DESC')
    users.each_with_index do |user, num|
      report.page.list(:list).add_row do |row|
        row.item('library').value(user.library.display_name) 
        row.item('role').value(user.required_role.display_name)
        row.item('full_name').value(user.agent.full_name)
        row.item('username').value(user.username)
        checkoutall_cnt = 0
        reserveall_cnt  = 0
        1.upto(12) do |month|
          yyyymm = "#{month > 3 ? options[:term] : options[:term].to_i + 1}#{"%02d" % month}"
          start_at = Time.zone.parse("#{yyyymm}01").beginning_of_month
          end_at   = Time.zone.parse("#{yyyymm}31").end_of_month
          checkout_cnt = user.checkouts.where("checked_at >= '#{start_at}' AND checked_at <= '#{end_at}'").count
          reserve_cnt  = user.reserves.where("created_at >= '#{start_at}' AND created_at <= '#{end_at}'").count
          row.item("checkout#{month}").value(checkout_cnt) 
          row.item("reserve#{month}").value(reserve_cnt)
          checkoutall_cnt = checkoutall_cnt + checkout_cnt
          reserveall_cnt  = reserveall_cnt  + reserve_cnt
        end
        row.item("checkoutall").value(checkoutall_cnt)
        row.item("reserveall").value(reserveall_cnt)
        # layout
        row.item('role_line').show unless user.required_role_id == users[num + 1].try(:required_role).try(:id) 
        unless user.library_id == users[num + 1].try(:library).try(:id)
          row.item('library_line').show 
          row.item('role_line').show
        end
        if user.library_id == users[num - 1].try(:library).try(:id)  
          row.item('library').hide if user.library_id == users[num - 1].try(:library).try(:id)
          row.item('role').hide if user.required_role_id == users[num - 1].try(:required_role).try(:id)
        end
      end
    end 
  end

  # 所属別利用統計 checkout: 122 / reserve: 133
  def self.set_departments_report_pdf(report, options = {})
    # term
    report.page.values(:term => options[:term])
    # list
    Department.all.each do |department| 
      report.page.list(:list).add_row do |row|
        row.item("department").value(department.display_name)
        checkoutall = 0
        reserveall  = 0
        1.upto(12) do |cnt|
          conditions = {
            :yyyymm        => "#{cnt > 3 ? options[:term] : options[:term].to_i + 1}#{"%02d" % cnt}",
            :department_id => department.id
          }
          checkout = Statistic.where(conditions.merge(:data_type => 122, :option => 1)).first.value rescue 0
          reserve  = Statistic.where(conditions.merge(:data_type => 133)).first.value rescue 0
          row.item("checkout#{cnt}").value(checkout)
          row.item("reserve#{cnt}").value(reserve)
          checkoutall = checkoutall + checkout
          reserveall  = reserveall  + reserve
        end
        row.item("checkoutall").value(checkoutall)
        row.item("reserveall").value(reserveall) 
      end
    end
  end

  # 年別利用統計 checkout: 121 / reserve: 133 
  def self.set_yearly_report_pdf(report, options = {})
    libraries = Library.real.all
    # term
    report.page.values(
      :year_start_at => options[:start_at], 
      :year_end_at   => options[:end_at]
    )
    # footer
    report.layout.config.list(:list) do
      events.on :footer_insert do |e|
        libraries.each_with_index do |library, num|
          conditions = "library_id = #{library.id} AND yyyymm >= #{options[:start_at]}04 AND yyyymm <= #{options[:end_at].to_i + 1}03" 
          checkout_all = Statistic.where(conditions + 'AND data_type = 121').sum(:value)
          reserve_all  = Statistic.where(conditions + 'AND data_type = 133').sum(:value)
          e.section.item("checkout_total##{num}").value(checkout_all)
          e.section.item("reserve_total##{num}").value(reserve_all)
          # footer layout
          targets = %w(library_footer_frame library_footer_column_line checkout_total reserve_total)
          targets.each { |target| e.section.item("#{target}##{num}").show }
        end
      end
    end
    # header
    libraries.each_with_index do |library, num|
      report.page.list(:list).header.item("library##{num}").value(library.display_name)
      # header layout
      targets = %w(library_header_frame library_header_column_line
        library_header_row_line library_header_checkout library_header_reserve)
      targets.each { |target| report.page.list(:list).header.item("#{target}##{num}").show }
    end
    # list data
    (options[:end_at].to_i - options[:start_at].to_i + 1).times do |cnt|
      report.page.list(:list).add_row do |row|
        year = options[:start_at].to_i + cnt
        row.item(:year).value(year)
        libraries.each_with_index do |library, num|
          conditions = "library_id = #{library.id} AND yyyymm >= #{year}04 AND yyyymm <= #{year + 1}03" 
          checkout = Statistic.where(conditions + 'AND data_type = 121').sum(:value)
          reserve  = Statistic.where(conditions + 'AND data_type = 133').sum(:value)
          row.item("checkout##{num}").value(checkout)
          row.item("reserve##{num}").value(reserve) 
          #layout
          targets = %w(library_detail_column_line1 library_detail_column_line2 library_detail_row_line checkout)
          targets.each { |target| row.item("#{target}##{num}").show }
        end
      end
    end
  end

  def self.get_monthly_report_pdf(term)
    libraries = Library.all
    checkout_types = CheckoutType.all
    user_groups = UserGroup.all
    begin 
      report = ThinReports::Report.new :layout => get_layout_path("monthly_report")

      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      report.start_new_page
      report.page.item(:date).value(Time.now)       
      report.page.item(:term).value(term)

      12.times do |t|
      end

      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_items")
        # items all libraries
        data_type = 111
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.items'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            end  
            row.item(:library_line).show if SystemConfiguration.get("statistic_report.monthly_public.items_not_use_checkout_types")
          end
          # items each checkout_types
          unless SystemConfiguration.get("statistic_report.monthly_public.items_not_use_checkout_types")
            checkout_types.each do |checkout_type|
              report.page.list(:list).add_row do |row|
                row.item(:type).value(I18n.t('statistic_report.items')) if libraries.size == 1 && checkout_types.first == checkout_type
                row.item(:option).value(checkout_type.display_name.localize)
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
                end 
                if checkout_type == checkout_types.last
                  row.item(:library_line).show
                  line_for_libraries(row)
                end
              end
            end
          end
=begin
        # missing items
        report.page.list(:list).add_row do |row|
          row.item(:library).value("(#{t('statistic_report.missing_items')})")
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :option => 1, :library_id => 0).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :option => 1, :library_id => 0).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            row.item(:library_line).show
          end  
        end
=end
        end
        # items each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.items')) if libraries.size == 1
            row.item(:library).value(library.display_name)
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            end  
            if SystemConfiguration.get("statistic_report.monthly_public.items_not_use_checkout_types")
              if library == libraries.last
                line(row)
              else
                row.item(:library_line).show
              end
            end
          end
          # items each checkout_types
          unless SystemConfiguration.get("statistic_report.monthly_public.items_not_use_checkout_types")
            checkout_types.each do |checkout_type|
              report.page.list(:list).add_row do |row|
                row.item(:option).value(checkout_type.display_name.localize)
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
                end  
                if checkout_type == checkout_types.last
                  row.item(:library_line).show
                  if library == libraries.last
                    line(row)
                  else
                    line_for_libraries(row)
                  end
                end
              end
            end
          end
=begin
          # missing items
          report.page.list(:list).add_row do |row|
            row.item(:library).value("(#{t('statistic_report.missing_items')})")
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :option => 1, :library_id => library.id).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :option => 1, :library_id => library.id).first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              row.item(:library_line).show
              line(row) if library == libraries.last
            end  
          end
=end
        end
      end

      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_open_days_of_libraries")
        # open days of each libraries
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.opens')) if libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 113, :library_id => library.id).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 113, :library_id => library.id).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end

      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_users_122")
        # checkout users all libraries
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkout_users'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => 0).no_condition.first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => 0).no_condition.first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show if SystemConfiguration.get("statistic_report.monthly_public.items_not_use_checkout_types")
          end
          unless SystemConfiguration.get("statistic_report.monthly_public.users_122_not_use_ages")
            # checkout users each user type
            5.downto(1) do |i|
              data_type = 122
              report.page.list(:list).add_row do |row|
                row.item(:option).value(I18n.t("statistic_report.user_type_#{i}"))
                sum = 0
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 0, user_type => i).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 0, user_type => i).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  sum = sum + value
                end  
                row.item("valueall").value(sum)
                if i == 1
                  row.item(:library_line).show
                  line_for_libraries(row)
                end
              end
            end
          end
        end
        # checkout users each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkout_users')) if libraries.size == 1
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => library.id).no_condition.first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => library.id).no_condition.first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            if SystemConfiguration.get("statistic_report.monthly_public.items_not_use_checkout_types")
              if library == libraries.last
                line(row)
              else
                row.item(:library_line).show
              end
            end
          end
          unless SystemConfiguration.get("statistic_report.monthly_public.users_122_not_use_ages")
            # checkout users each user type
            5.downto(1) do |i|
              data_type = 122
              report.page.list(:list).add_row do |row|
                row.item(:option).value(I18n.t("statistic_report.user_type_#{i}"))
                sum = 0
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 0, user_type => i).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 0, user_type => i).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  sum = sum + value
                end  
                row.item("valueall").value(sum)
                if i == 1
                  row.item(:library_line).show
                  if library == libraries.last
                    line(row) if library == libraries.last
                  else
                    line_for_libraries(row)
                  end
                end
              end
            end  
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_daily_average_of_checkout")
        # daily average of checkout users all library
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.average_checkout_users'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => 0, :option => 4).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => 0, :option => 4).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum/12)
            row.item(:library_line).show
          end
        end
        # daily average of checkout users each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.average_checkout_users')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => library.id, :option => 4).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => library.id, :option => 4).first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum/12)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_checkout_items")
        # checkout items all libraries
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkout_items'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0).no_condition.first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0).no_condition.first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show if SystemConfiguration.get("statistic_report.monthly_public.checkout_items_not_use_types")
          end
          unless SystemConfiguration.get("statistic_report.monthly_public.checkout_items_not_use_types")
            3.times do |i|
              report.page.list(:list).add_row do |row|
                row.item(:type).value(I18n.t('statistic_report.checkout_items')) if libraries.size == 1 && i == 1
                row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
                sum = 0
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :option => i+1, :age => nil).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :option => i+1, :age => nil).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  sum = sum + value
                end  
                row.item("valueall").value(sum)
                if i == 2
                  row.item(:library_line).show
                  line_for_libraries(row)
                end
              end
            end
          end
        end

        # checkout items each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkout_items')) if libraries.size == 1
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id).no_condition.first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id).no_condition.first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            if SystemConfiguration.get("statistic_report.monthly_public.checkout_items_not_use_types")
              if library == libraries.last
                line(row)
              else
                row.item(:library_line).show
              end
            end
          end
          unless SystemConfiguration.get("statistic_report.monthly_public.checkout_items_not_use_types")
            3.times do |i|
              report.page.list(:list).add_row do |row|
                row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
                sum = 0
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id, :option => i+1, :age => nil).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id, :option => i+1, :age => nil).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  sum = sum + value
                end  
                row.item("valueall").value(sum)
                if i == 2
                  row.item(:library_line).show
                  if library == libraries.last
                    line(row)
                  else
                    line_for_libraries(row)
                  end
                end
              end
            end
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.checkout_items_not_use_types")
        # checkout items each user_group
        if libraries.size > 1
          user_groups.each do |user_group|
            report.page.list(:list).add_row do |row|
              if user_group == user_groups.first
                row.item(:type).value(I18n.t('statistic_report.checkout_items_each_user_groups'))
                row.item(:library).value(I18n.t('statistic_report.all_library')) 
              end
              row.item(:option).value(user_group.display_name.localize)   
              sum = 0
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :user_group_id => user_group.id).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :user_group_id => user_group.id).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                sum = sum + value
              end  
              row.item("valueall").value(sum)
              if user_group == user_groups.last
                row.item(:library_line).show 
                line_for_libraries(row)
              end
            end
          end
        end
        libraries.each do |library|
          user_groups.each do |user_group|
            report.page.list(:list).add_row do |row|
              row.item(:type).value(I18n.t('statistic_report.checkout_items_each_user_groups')) if libraries.size == 1 && libraries.first == library && user_groups.first == user_group
              row.item(:library).value(library.display_name.localize) if user_group == user_groups.first
              row.item(:option).value(user_group.display_name.localize)
              sum = 0
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id).no_condition.first.value rescue 0 
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id).no_condition.first.value rescue 0 
                end
                row.item("value#{t+1}").value(to_format(value))
                sum = sum + value
              end  
              row.item("valueall").value(sum)
              if user_group == user_groups.last
	        row.item(:library_line).show
                if library == libraries.last
                  line(row)
                else
                  line_for_libraries(row)
                end
              end
            end  
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_daily_average_of_checkout_items")
        # daily average of checkout items all library
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.average_checkout_items'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :option => 4).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :option => 4).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum/12)
            row.item(:library_line).show
          end
        end
        # daily average of checkout items each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.average_checkout_items')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id, :option => 4).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id, :option => 4).first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum/12)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_remind_checkout_items")
        # reminder checkout items
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.remind_checkouts'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :option => 5).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :option => 5).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show
          end
        end	
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.remind_checkouts')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name.localize)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id, :option => 5).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library_id, :option => 5).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_checkin_items")
        # checkin items
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkin_items'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0).no_condition.first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0).no_condition.first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show
          end
        end
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkin_items')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library.id).no_condition.first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library.id).no_condition.first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
	    row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_daily_average_of_checkin_items")
        # daily average of checkin items all library
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.average_checkin_items'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0, :option => 4).first.value rescue 0
              else
                 value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0, :option => 4).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum/12)
            row.item(:library_line).show
          end
        end 

        # daily average of checkin items each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.average_checkin_items')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library.id, :option => 4).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library.id, :option => 4).first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum/12)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_checkin_items_remindered")
        # checkin items remindered
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkin_remindered'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0, :option => 5).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0, :option => 5).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show
          end
        end 
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkin_remindered')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name.localize)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library.id, :option => 5).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library_id, :option => 5).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_users_112")
        # all users all libraries
        data_type = 112
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.users'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
              unless SystemConfiguration.get("statistic_report.monthly_public.users_not_use_types")
            row.item(:option).value(I18n.t('statistic_report.all_users'))
            end
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              row.item(:library_line).show if SystemConfiguration.get("statistic_report.monthly_public.users_not_use_types")
            end  
          end
          unless SystemConfiguration.get("statistic_report.monthly_public.users_112_not_use_types")
            # users each user type
            5.downto(1) do |i|
              report.page.list(:list).add_row do |row|
                row.item(:option).value(I18n.t("statistic_report.user_type_#{i}"))
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 0, :user_type => i).first.value rescue 0
                  else	
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 0, :user_type => i).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
                end
              end  
            end
            # unlocked users all libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.unlocked_users'))
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 1).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 1).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              end  
            end
            # locked users all libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.locked_users'))
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 2).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 2).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              end  
            end
            # provisional users all libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.user_provisional'))
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 3).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 3).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
                line_for_libraries(row)
              end
            end
          end
        end

        # users each library
        libraries.each do |library|
          # all users
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.users')) if libraries.size == 1
            row.item(:library).value(library.display_name)
            unless SystemConfiguration.get("statistic_report.monthly_public.users_112_not_use_types")
              row.item(:option).value(I18n.t('statistic_report.all_users'))
            end
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            end  
            if SystemConfiguration.get("statistic_report.monthly_public.users_112_not_use_types")
              if library == libraries.last 
                line(row)
              else
                row.item(:library_line).show
              end
            end
          end
          unless SystemConfiguration.get("statistic_report.monthly_public.users_112_not_use_types")
            # users each user type
            5.downto(1) do |i|
              report.page.list(:list).add_row do |row|
                row.item(:option).value(I18n.t("statistic_report.user_type_#{i}"))
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 0, :user_type => i).first.value rescue 0 
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 0, :user_type => i).first.value rescue 0 
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
                end
              end  
            end
            # unlocked users
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.unlocked_users'))
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 1).first.value rescue 0 
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 1).first.value rescue 0 
                end
                row.item("value#{t+1}").value(to_format(value))
                row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              end  
            end
            # locked users
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.locked_users'))
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 2).first.value rescue 0 
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 2).first.value rescue 0 
                end
                row.item("value#{t+1}").value(to_format(value))
                row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              end  
            end
            # provisional users all libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.user_provisional'))
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 3).first.value rescue 0 
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 3).first.value rescue 0 
                end
                row.item("value#{t+1}").value(to_format(value))
                row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              end  
              row.item(:library_line).show
              if library == libraries.last
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_reserves")
        # reserves all libraries
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.reserves'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => 0).no_condition.first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => 0).no_condition.first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show if SystemConfiguration.get("statistic_report.monthly_public.reserves_not_use_types")
          end
          unless SystemConfiguration.get("statistic_report.monthly_public.reserves_not_use_types")
            # reserves on counter all libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.on_counter'))
              sum = 0
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => 0, :option => 1, :age => nil).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => 0, :option => 1, :age => nil).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                sum = sum + value
              end
              row.item("valueall").value(sum)
            end
            # reserves from OPAC all libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.from_opac'))
              sum = 0
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => 0, :option => 2, :age => nil).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => 0, :option => 2, :age => nil).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                sum = sum + value
              end
              row.item("valueall").value(sum)
              line_for_libraries(row)
            end
          end
        end
        # reserves each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.reserves')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => library.id).no_condition.first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => library.id).no_condition.first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            if SystemConfiguration.get("statistic_report.monthly_public.reserves_not_use_types")
              if library == libraries.last
                line(row)
              else
                row.item(:library_line).show
              end
            end
          end
          unless SystemConfiguration.get("statistic_report.monthly_public.reserves_not_use_types")
            # reserves on counter each libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.on_counter'))
              sum = 0
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => library.id, :option => 1, :age => nil).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => library.id, :option => 1, :age => nil).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                sum = sum + value
              end
              row.item("valueall").value(sum)
            end
            # reserves from OPAC each libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.from_opac'))
              sum = 0
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => library.id, :option => 2, :age => nil).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => library.id, :option => 2, :age => nil).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                sum = sum + value
              end
              row.item("valueall").value(sum)
              row.item(:library_line).show
              if library == libraries.last
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_questions")
        # questions all libraries
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.questions'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 143, :library_id => 0).no_condition.first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 143, :library_id => 0).no_condition.first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show
          end
        end
        # questions each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.questions')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 143, :library_id => library.id).no_condition.first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 143, :library_id => library.id).no_condition.first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_visiters")
        # visiters all libraries
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.visiters'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 116, :library_id => 0).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 116, :library_id => 0).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item(:library_line).show
          end
        end
        # visiters of each libraries
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.visiters')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 116, :library_id => library.id).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 116, :library_id => library.id).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_consultations")
        # consultations all libraries
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.consultations'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 114, :library_id => 0).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 114, :library_id => 0).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item(:library_line).show
          end
        end
        # consultations of each libraries
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.consultations')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 114, :library_id => library.id).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 114, :library_id => library.id).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_copies")
        # copies all libraries
        if libraries.size > 1
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.copies'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 115, :library_id => 0).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 115, :library_id => 0).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item(:library_line).show
          end
        end
        # copies of each libraries
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.copies')) if libraries.size == 1 && libraries.first == library
            row.item(:library).value(library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 115, :library_id => library.id).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 115, :library_id => library.id).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item(:library_line).show
            line(row) if library == libraries.last
          end
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      logger.error $@.join('\n')
      return false
    end	
  end

  def self.get_monthly_report_tsv(term)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_monthly_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:library, 'statistic_report.library'],
      [:option, 'statistic_report.option']
    ]
    libraries = Library.all
    checkout_types = CheckoutType.all
    user_groups = UserGroup.all
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print"\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      9.times do |t|
        row << I18n.t('statistic_report.month', :num => t+4)
        columns << ["#{term}#{"%02d" % (t + 4)}"]
      end
      3.times do |t|
        row << I18n.t('statistic_report.month', :num => t+1)
        columns << ["#{term.to_i + 1}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_items")
        # items all libraries
        data_type = 111
        if libraries.size > 1
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.items')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row << to_format(value)
            end
          end  
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.items_not_use_checkout_types")
            # items each checkout_types
            checkout_types.each do |checkout_type|
              row = []
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.items')
                when :library
                  row << I18n.t('statistic_report.all_library')
                when :option
                  row << checkout_type.display_name.localize
                when "sum"
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
                  row << to_format(value)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
                  row << to_format(value)
                end
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
        end
        # items each library
        libraries.each do |library|
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.items')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
              row << to_format(value)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.items_not_use_checkout_types")
            # items each checkout_types
            checkout_types.each do |checkout_type|
              row = []
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.items')
                when :library
                  row << library.display_name
                when :option
                  row << checkout_type.display_name.localize
                when "sum"
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
                  row << to_format(value)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
                  row << to_format(value)
                end
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_open_days_of_libraries")
        # open days of each libraries
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.opens')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 113, :library_id => library.id).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_users_122")
        # checkout users all libraries
        data_type = 122
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_users')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.users_122_not_use_ages")
            # checkout users each user type
            5.downto(1) do |i|
              sum = 0
              row = []
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.checkout_users')
                when :library
                  row << I18n.t('statistic_report.all_library')
                when :option
                  row << I18n.t("statistic_report.user_type_#{i}")
                when "sum"
                  row << to_format(sum)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :option => 0, user_type => i).first.value rescue 0
                  sum += value
                  row << to_format(value)
                end  
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
          # checkout users each library
          libraries.each do |library|
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.checkout_users')
              when :library
                row << library.display_name
              when :option
                row << ""
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
                sum += value
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
            unless SystemConfiguration.get("statistic_report.monthly_public.users_122_not_use_ages")
              # checkout users each user type
              5.downto(1) do |i|
                sum = 0
                row = []
                columns.each do |column|
                  case column[0]
                  when :type
                    row << I18n.t('statistic_report.checkout_users')
                  when :library
                    row << library.display_name
                  when :option
                    row << I18n.t("statistic_report.user_type_#{i}")
                  when "sum"
                    row << to_format(sum)
                  else
                    value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :option => 0, user_type => i).first.value rescue 0
                    sum += value
                    row << to_format(value)
                  end  
                end
                output.print "\""+row.join("\"\t\"")+"\"\n"
              end
            end
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_daily_average_of_checkout")
        # daily average of checkout users all library
        data_type = 122
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.average_checkout_users')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum/12)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :option => 4).first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # daily average of checkout users each library
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.average_checkout_users')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum/12)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :option => 4).first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_checkout_items")
        # checkout items all libraries
        data_type = 121
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_items')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.checkout_items_not_use_types")
            # checkout items all libraries each item types
            3.times do |i|
              sum = 0
              row = []
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.checkout_items')
                when :library
                  row << I18n.t('statistic_report.all_library')
                when :option
                  row << I18n.t("statistic_report.item_type_#{i+1}")
                when "sum"
                  row << to_format(sum)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :option => i+1, :age => nil).first.value rescue 0
                  sum += value
                  row << to_format(value)
                end  
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
        end
        # checkout items each library
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_items')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 121, :library_id => library.id).no_condition.first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.checkout_items_not_use_types")
            3.times do |i|
              sum = 0
              row = []
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.checkout_items')
                when :library
                  row << library.display_name
                when :option
                  row << I18n.t("statistic_report.item_type_#{i+1}")
                when "sum"
                  row << to_format(sum)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :option => i+1, :age => nil).first.value rescue 0
                  sum += value
                  row << to_format(value)
                end  
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.checkout_items_not_use_types")
        # checkout items each user_group
        if libraries.size > 1
          user_groups.each do |user_group|
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.checkout_items_each_user_groups')
              when :library
                row << I18n.t('statistic_report.all_library')
              when :option
                row << user_group.display_name.localize
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :user_group_id => user_group.id).first.value rescue 0
                sum += value
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
        libraries.each do |library|
          user_groups.each do |user_group|
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.checkout_items_each_user_groups')
              when :library
                row << library.display_name.localize
              when :option
                row << user_group.display_name.localize
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
                sum += value
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_daily_average_of_checkout_items")
        # daily average of checkout items all library
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.average_checkout_items')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum/12) rescue 0
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 121, :library_id => 0, :option => 4).first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # daily average of checkout items each library
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.average_checkout_items')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum/12)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 121, :library_id => library.id, :option => 4).first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_remind_checkout_items")
        # remind checkout items
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.remind_checkouts')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        libraries.each do |library|     
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.remind_checkouts')
            when :library
              row << library.display_name.localize
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :option => 5).first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_checkin_items")
        # checkin items
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkin_items')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 151, :library_id => 0).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkin_items')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 151, :library_id => library.id).no_condition.first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_daily_average_of_checkin_items")
        # daily average of checkin items all library
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.average_checkin_items')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum/12)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 151, :library_id => 0, :option => 4).first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # daily average of checkin items each library
        libraries.each do |library|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.average_checkin_items')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum/12)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 151, :library_id => library.id, :option => 4).first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_checkin_items_remindered")
        # checkin items remindered
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkin_remindered')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 151, :library_id => 0, :option => 5).first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        libraries.each do |library|     
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkin_remindered')
            when :library
              row << library.display_name.localize
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 151, :library_id => library.id, :option => 5).first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_users_112")
        # all users all libraries
        if libraries.size > 1
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.users')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << I18n.t('statistic_report.all_users')
            when "sum"
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => 0).no_condition.first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => 0).no_condition.first.value rescue 0
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.users_112_not_use_types")
            # users each user type
            5.downto(1) do |i|
              row = []
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.users')
                when :library
                  row << I18n.t('statistic_report.all_library')
                when :option
                  row << I18n.t("statistic_report.user_type_#{i}")
                when "sum"
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => 0, :option => 0, :user_type => i).first.value rescue 0
                  row << to_format(value)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => 0, :option => 0, :user_type => i).first.value rescue 0
                  row << to_format(value)
                end  
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
            # unlocked users all libraries
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.users')
              when :library
                row << I18n.t('statistic_report.all_library')
              when :option
                row << I18n.t('statistic_report.unlocked_users')
              when "sum"
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => 0, :option => 1).first.value rescue 0
                row << to_format(value)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => 0, :option => 1).first.value rescue 0
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
            # locked users all libraries
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.users')
              when :library
                row << I18n.t('statistic_report.all_library')
              when :option
                row << I18n.t('statistic_report.locked_users')
              when "sum"
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => 0, :option => 2).first.value rescue 0
                row << to_format(value)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => 0, :option => 2).first.value rescue 0
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
            # provisional users all libraries
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.users')
              when :library
                row << I18n.t('statistic_report.all_library')
              when :option
                row << I18n.t('statistic_report.user_provisional')
              when "sum"
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => 0, :option => 3).first.value rescue 0
                row << to_format(value)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => 0, :option => 3).first.value rescue 0
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
        # users each library
        libraries.each do |library|
          # all users
          row = []
          columns.each do |column|
           case column[0]
            when :type
              row << I18n.t('statistic_report.users')
            when :library
              row << library.display_name.localize
            when :option
              row << I18n.t('statistic_report.user_provisional')
            when "sum"
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => library.id).no_condition.first.value rescue 0 
              row << to_format(value)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => library.id).no_condition.first.value rescue 0 
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.users_112_not_use_types")
            # users each user type
            5.downto(1) do |i|
              row = []
              columns.each do |column|
               case column[0]
                when :type
                  row << I18n.t('statistic_report.users')  
                when :library
                  row << library.display_name.localize
                when :option
                  row << I18n.t("statistic_report.user_type_#{i}")
                when "sum"
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => library.id, :option => 0, :user_type => i).first.value rescue 0 
                  row << to_format(value)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => library.id, :option => 0, :user_type => i).first.value rescue 0 
                  row << to_format(value)
                end  
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
            # unlocked users
            row = []
            columns.each do |column|
             case column[0]
              when :type
                row << I18n.t('statistic_report.users')  
              when :library
                row << library.display_name.localize
              when :option
                row << I18n.t('statistic_report.unlocked_users')
              when "sum"
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => library.id, :option => 1).first.value rescue 0 
                row << to_format(value)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => library.id, :option => 1).first.value rescue 0 
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
            # locked users
            row = []
            columns.each do |column|
             case column[0]
              when :type
                row << I18n.t('statistic_report.users')  
              when :library
                row << library.display_name.localize
              when :option
                row << I18n.t('statistic_report.locked_users')
              when "sum"
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => library.id, :option => 2).first.value rescue 0 
                row << to_format(value)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => library.id, :option => 2).first.value rescue 0 
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
            # provisional users all libraries
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.users')  
              when :library
                row << library.display_name.localize
              when :option
                row << I18n.t('statistic_report.user_provisional')
              when "sum"
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => library.id, :option => 3).first.value rescue 0 
                row << to_format(value)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => library.id, :option => 3).first.value rescue 0 
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_reserves")
        # reserves all libraries
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.reserves')  
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 133, :library_id => 0).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.reserves_not_use_types")
            # reserves on counter all libraries
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.reserves')  
              when :library
                row << I18n.t('statistic_report.all_library')
              when :option
                row << I18n.t('statistic_report.on_counter')
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 133, :library_id => 0, :option => 1, :age => nil).first.value rescue 0
                sum += value
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
            # reserves from OPAC all libraries
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.reserves')  
              when :library
                row << I18n.t('statistic_report.all_library')
              when :option
                row << I18n.t('statistic_report.from_opac')
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 133, :library_id => 0, :option => 2, :age => nil).first.value rescue 0
                sum += value
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
        # reserves each library
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.reserves')  
            when :library
              row << library.display_name.localize
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 133, :library_id => library.id).no_condition.first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.monthly_public.reserves_not_use_types")
            # reserves on counter each libraries
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.reserves')  
              when :library
                row << library.display_name.localize
              when :option
                row << I18n.t('statistic_report.on_counter')
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 133, :library_id => library.id, :option => 1, :age => nil).first.value rescue 0
                sum += value
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
            # reserves from OPAC each libraries
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.reserves')  
              when :library
                row << library.display_name.localize
              when :option
                row << I18n.t('statistic_report.from_opac')
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 133, :library_id => library.id, :option => 2, :age => nil).first.value rescue 0
                sum += value
                row << to_format(value)
              end  
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_questions")
        # questions all libraries
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.questions')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 143, :library_id => 0).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # questions each library
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.questions')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 143, :library_id => library.id).no_condition.first.value rescue 0 
              sum += value
              row << to_format(value)
            end
          end  
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_visiters")
        # visiters all libraries
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.visiters')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 116, :library_id => 0).first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # visiters of each libraries
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.visiters')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 116, :library_id => library.id).first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_consultations")
        # consultations all libraries
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.consultations')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 114, :library_id => 0).first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # consultations of each libraries
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.consultations')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 114, :library_id => library.id).first.value rescue 0 
              sum = value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.monthly_public.not_use_copies")
        # copies all libraries
        if libraries.size > 1
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.copies')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 115, :library_id => 0).first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # copies of each libraries
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.copies')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 115, :library_id => library.id).first.value rescue 0 
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
    end
    return tsv_file
  end

  def self.get_daily_report_pdf(term)
    libraries = Library.all
    logger.error "create daily statistic report: #{term}"

    begin
      report = ThinReports::Report.new :layout => get_layout_path("daily_report")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      num_for_last_page = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i - 26
      [1,14,27].each do |start_date| # for 3 pages
        report.start_new_page
        report.page.item(:date).value(Time.now)
        report.page.item(:year).value(term[0,4])
        report.page.item(:month).value(term[4,6])        
        # header
        if start_date != 27
          13.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
        else
          num_for_last_page.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
          report.page.list(:list).header.item("column#13").value(I18n.t('statistic_report.sum'))
        end

        unless SystemConfiguration.get("statistic_report.daily_report.not_use_items")
          # items all libraries
          data_type = 211
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.items'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
              end
            end
            row.item(:library_line).show
          end
          # items each libraries
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
              row.item(:library).value(library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 211, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 211, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
                end
              end
              row.item(:library_line).show
              line(row) if library == libraries.last
            end
          end
        end
        unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkout_users")
          # checkout users all libraries
          data_type = 222
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkout_users'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0).no_condition
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
            row.item(:library_line).show if SystemConfiguration.get("statistic_report.daily_report.checkout_users_not_use_types")
          end
          # each user type
          unless SystemConfiguration.get("statistic_report.daily_report.checkout_users_not_use_types")
            5.downto(1) do |type|
              report.page.list(:list).add_row do |row|
                row.item(:option).value(I18n.t("statistic_report.user_type_#{type}"))
                if start_date != 27
                  13.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 0, :user_type => type).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                  end
                else
                  num_for_last_page.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 0, :user_type => type).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                  end
                  sum = 0
                  datas = Statistic.where(["yyyymm = ? AND data_type = ? AND library_id = ? AND option = 0 AND user_type = ? ", term, data_type, 0, type])
                  datas.each do |data|
                    sum += data.value                               
                  end
                  row.item("value#13").value(sum)
                end
                if type == 1
                  line_for_libraries(row)
                end
              end
            end
          end
          # checkout users each libraries
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
            row.item(:library).value(library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id).no_condition
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              if SystemConfiguration.get("statistic_report.daily_report.checkout_users_not_use_types")
                if library == libraries.last
                  line(row)
                else
                  row.item(:library_line).show
                end
              end
            end
            unless SystemConfiguration.get("statistic_report.daily_report.checkout_users_not_use_types")
              # each user type
              5.downto(1) do |type|
                report.page.list(:list).add_row do |row|
                  row.item(:option).value(I18n.t("statistic_report.user_type_#{type}"))
                  if start_date != 27
                    13.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option =>0, :user_type => type).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                    end
                  else
                    num_for_last_page.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 0, user_type => type).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                    end
                    sum = 0
                    datas = Statistic.where(["yyyymm = ? AND data_type = ? AND library_id = ? AND option = 0 AND user_type = ?", term, data_type, library.id, type])
                    datas.each do |data|
                      sum += data.value                               
                    end
                    row.item("value#13").value(sum)
                  end
                  if type == 1
                    if library == libraries.last
                      line(row) if library == libraries.last
                    else
                      line_for_libraries(row)
                    end
                  end 
                end
              end
            end
          end
        end

        unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkout_items")
          # checkout items all libraries
          data_type = 221
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkout_items'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_conditionfirst.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0).no_condition
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
            row.item(:library_line).show if SystemConfiguration.get("statistic_report.daily_report.checkout_items_not_use_types")
          end
          unless SystemConfiguration.get("statistic_report.daily_report.checkout_items_not_use_types")
            3.times do |i|
              report.page.list(:list).add_row do |row|
                row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
                if start_date != 27
                  13.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => i+1).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                  end
                else
                  num_for_last_page.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => i+1).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                  end
                  sum = 0
                  datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :option => i+1)
                  datas.each do |data|
                    sum = sum + data.value
                  end
                  row.item("value#13").value(sum)
                end
                line_for_libraries(row) if i == 2
              end
            end
          end
        
          # checkout items each libraries
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
              row.item(:library).value(library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id).no_condition
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              if SystemConfiguration.get("statistic_report.daily_report.checkout_items_not_use_types")
                if library == libraries.last
                  line(row)
                else
                  row.item(:library_line).show
                end
              end
            end
            unless SystemConfiguration.get("statistic_report.daily_report.checkout_items_not_use_types")
              3.times do |i|
                report.page.list(:list).add_row do |row|
                  row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
                  if start_date != 27
                    13.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => i+1).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                    end
                  else
                    num_for_last_page.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => i+1).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                    end
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :option => i+1)
                    datas.each do |data|
                      sum = sum + data.value
                    end
                    row.item("value#13").value(sum)
                  end
                  if i == 2
                    if library == libraries.last
                      line(row)
                    else
                      line_for_libraries(row)
                    end
                  end
                end
              end
            end
          end
        end

        unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkout_items_reminded")
          # remind checkout items
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.remind_checkouts'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :option => 5)
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
            row.item(:library_line).show
          end
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
              row.item(:library).value(library.display_name.localize)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 5).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 5).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :option => 5)
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              row.item(:library_line).show
              line(row) if library == libraries.last
            end
          end
        end
    
        unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkin_items")
          # checkin items
          data_type = 251
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkin_items'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0).no_condition
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
            row.item(:library_line).show
          end
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
              row.item(:library).value(library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id).no_condition
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              row.item(:library_line).show
              line(row) if library == libraries.last
            end
          end
        end

        unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkin_items_reminded")
          # checkin items remindered
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkin_remindered'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0).no_condition
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
            row.item(:library_line).show
          end
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
              row.item(:library).value(library.display_name.localize)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 5).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 5).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :option => 5)
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              row.item(:library_line).show
              line(row) if library == libraries.last
            end
          end
        end

        unless SystemConfiguration.get("statistic_report.daily_report.not_use_reserves")
          # reserves all libraries
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.reserves'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else  
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(:yyyymm => term, :data_type => 233, :library_id => 0).no_condition
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
            row.item(:library_line).show if SystemConfiguration.get("statistic_report.daily_report.reserves_not_use_types")
          end
          unless SystemConfiguration.get("statistic_report.daily_report.reserves_not_use_types")
            # reserves on counter all libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.on_counter'))
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => 0, :option => 1).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else  
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => 0, :option => 1).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => 233, :library_id => 0, :option => 1)
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
            end
            # reserves from OPAC all libraries
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.from_opac'))
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => 0, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else  
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => 0, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => 233, :library_id => 0, :option => 2)
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              line_for_libraries(row)
            end
          end
          # reserves each library
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
              row.item(:library).value(library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else  
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => 233, :library_id => library.id).no_condition
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              if SystemConfiguration.get("statistic_report.daily_report.reserves_not_use_types")
                if library == libraries.last
                  line(row)
                else
                  row.item(:library_line).show
                end
              end
            end
          end
          unless SystemConfiguration.get("statistic_report.daily_report.reserves_not_use_types")
            # on counter
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.on_counter'))
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => library.id, :option => 1).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else  
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => library.id, :option => 1).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => 233, :library_id => library.id, :option => 1)
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
            end
            # from OPAC
            report.page.list(:list).add_row do |row|
              row.item(:option).value(I18n.t('statistic_report.from_opac'))
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => library.id, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else  
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => library.id, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => 233, :library_id => library.id, :option => 2)
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              if library == libraries.last
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end
        end
       
        unless SystemConfiguration.get("statistic_report.daily_report.not_use_questions")
          # questions all libraries
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.questions'))
            row.item(:library).value(I18n.t('statistic_report.all_library'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 243, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end  
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 243, :library_id => 0).no_condition.first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0              
              datas = Statistic.where(:yyyymm => term, :data_type => 243, :library_id => 0).no_condition
              datas.each do |data|
                sum = sum + data.value 
              end
              row.item("value#13").value(sum)
            end
            row.item(:library_line).show
          end
          # questions each library
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
              row.item(:library).value(library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 243, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end  
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 243, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => 243, :library_id => library.id).no_condition
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              row.item(:library_line).show
              line(row) if library == libraries.last
            end
          end
        end

        unless SystemConfiguration.get("statistic_report.daily_report.not_use_consultations")
          # consultations each library
          libraries.each do |library|
            report.page.list(:list).add_row do |row|
              row.item(:type).value(I18n.t('statistic_report.consultations')) if libraries.first == library
              row.item(:library).value(library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 214, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end  
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 214, :library_id => library.id).no_condition.first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => 214, :library_id => library.id).no_condition
                datas.each do |data|
                  sum = sum + data.value
                end
                row.item("value#13").value(sum)
              end
              row.item(:library_line).show
              line(row) if library == libraries.last
            end
          end
        end
      end
      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_daily_report_tsv(term)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_daily_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    libraries = Library.all
    days = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:library, 'statistic_report.library'],
      [:option, 'statistic_report.option']
    ]
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      days.times do |t|
        row << I18n.t('statistic_report.date', :num => t+1)
        columns << ["#{term}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      unless SystemConfiguration.get("statistic_report.daily_report.not_use_items")
        # items all libraries
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.items')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            value = Statistic.where(:yyyymmdd => "#{term}#{days}", :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
            row << to_format(value)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
            row << to_format(value)
          end
        end  
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # items each libraries
        libraries.each do |library|
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.items')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              value = Statistic.where(:yyyymmdd => "#{term}#{days}", :data_type => 211, :library_id => library.id).no_condition.first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => library.id).no_condition.first.value rescue 0
              row << to_format(value)
            end
          end  
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkout_users")
        # checkout users all libraries
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 222, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end  
        output.print "\""+row.join("\"\t\"")+"\"\n"
        unless SystemConfiguration.get("statistic_report.daily_report.checkout_users_not_use_types")
          # each user type
          5.downto(1) do |type|
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.checkout_users')
              when :library
                row << I18n.t('statistic_report.all_library')
              when :option
                row << I18n.t("statistic_report.user_type_#{type}")
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymmdd => column[0], :data_type => 222, :library_id => 0, :option => 0, :user_type => type).first.value rescue 0
                sum += value
                row << to_format(value)
              end
            end
          end  
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # checkout users each libraries
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_users')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 222, :library_id => library.id).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end  
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.daily_report.checkout_users_not_use_types")
            # each user type
            5.downto(1) do |type|
              sum = 0
              row = []
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.checkout_users')
                when :library
                  row << library.display_name
                when :option
                  row << I18n.t("statistic_report.user_type_#{type}")
                when "sum"
                  row << to_format(sum)
                else
                  value = Statistic.where(:yyyymmdd => column[0], :data_type => 222, :library_id => library.id, :option =>0, :user_type => type).first.value rescue 0
                  sum += value
                  row << to_format(value)
                end
              end
            end  
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
      end
      unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkout_items")
        # checkout items all libraries
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"  
        unless SystemConfiguration.get("statistic_report.daily_report.checkout_items_not_use_types")
          3.times do |i|
            sum = 0
            row = []
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.checkout_items')
              when :library
                row << I18n.t('statistic_report.all_library')
              when :option
                row << I18n.t("statistic_report.item_type_#{i+1}")
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :library_id => 0, :option => i+1).first.value rescue 0
                sum += value
                row << to_format(value)
              end
            end 
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end        
        # checkout items each libraries
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_items')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :library_id => library.id).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless SystemConfiguration.get("statistic_report.daily_report.checkout_items_not_use_types")
            3.times do |i|
              sum = 0
              row = []
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.checkout_items')
                when :library
                  row << library.display_name
                when :option
                  row << I18n.t("statistic_report.item_type_#{i+1}")
                when "sum"
                  row << to_format(sum)
                else
                  value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :library_id => library.id, :option => i+1).first.value rescue 0
                  sum += value
                  row << to_format(value)
                end
              end 
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end  
        end
      end
      unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkout_items_reminded")
        # checkout items reminded
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.remind_checkouts')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :library_id => 0, :option => 5).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"  
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.remind_checkouts')
            when :library
              row << library.display_name.localize
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :library_id => library.id, :option => 5).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"  
        end
      end

      unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkin_items")
        # checkin items
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkin_items')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 251, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end 
        output.print "\""+row.join("\"\t\"")+"\"\n"
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkin_items')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 251, :library_id => library.id).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end

      unless SystemConfiguration.get("statistic_report.daily_report.not_use_checkin_items_reminded")
        # checkin items reminded
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkin_remindered')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 251, :library_id => 0, :option => 5).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"  
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkin_remindered')
            when :library
              row << library.display_name.localize
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 251, :library_id => library.id, :option => 5).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"  
        end
      end

      unless SystemConfiguration.get("statistic_report.daily_report.not_use_reserves")
        # reserves all libraries
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 233, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end 
        output.print "\""+row.join("\"\t\"")+"\"\n"
        unless SystemConfiguration.get("statistic_report.daily_report.reserves_not_use_types")
          # reserves on counter all libraries
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.reserves')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << I18n.t('statistic_report.on_counter')
            when "sum"
             row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 233, :library_id => 0, :option => 1).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
          # reserves from OPAC all libraries
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.reserves')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << I18n.t('statistic_report.from_opac')
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 233, :library_id => 0, :option => 2).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # reserves each library
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.reserves')
            when :library
              row << library.display_name.localize
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 233, :library_id => library.id).no_condition.first.value rescue 0
	      sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        unless SystemConfiguration.get("statistic_report.daily_report.reserves_not_use_types")
          # on counter
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.reserves')
            when :library
              row << library.display_name.localize
            when :option
              row << I18n.t('statistic_report.on_counter')
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 233, :library_id => library.id, :option => 1).first.value rescue 0
	      sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
          # from OPAC
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.reserves')
            when :library
              row << library.display_name.localize
            when :option
              row << I18n.t('statistic_report.from_opac')
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 233, :library_id => library.id, :option => 2).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.daily_report.not_use_questions")
        # questions all libraries
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.questions')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 243, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end 
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # questions each library
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.questions')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 243, :library_id => library.id).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      unless SystemConfiguration.get("statistic_report.daily_report.not_use_consultations")
        # consultations each library
        libraries.each do |library|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.consultations')
            when :library
              row << library.display_name
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 214, :library_id => library.id).no_condition.first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end 
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
    end
    return tsv_file
  end

  def self.get_timezone_report_pdf(start_at, end_at)
    #default setting 9 - 20
    open = SystemConfiguration.get("statistic_report.open")
    hours = SystemConfiguration.get("statistic_report.hours")

    libraries = Library.all
    logger.error "create daily timezone report: #{start_at} - #{end_at}"

    begin
      report = ThinReports::Report.new :layout => get_layout_path("timezone_report")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      report.start_new_page
      report.page.item(:date).value(Time.now)
      report.page.item(:year_start_at).value(start_at[0,4])
      report.page.item(:month_start_at).value(start_at[4,2])
      report.page.item(:date_start_at).value(start_at[6,2])
      report.page.item(:year_end_at).value(end_at[0,4])
      report.page.item(:month_end_at).value(end_at[4,2])
      report.page.item(:date_end_at).value(end_at[6,2]) rescue nil 

      # header 
      hours.times do |t|
        report.page.list(:list).header.item("column##{t+1}").value("#{t+open}#{I18n.t('statistic_report.hour')}")
      end
      report.page.list(:list).header.item("column#15").value(I18n.t('statistic_report.sum'))

      # checkout users all libraries
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.checkout_users'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        hours.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 322 AND library_id = ? AND hour = ?", 0, t+open]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value##{t+1}").value(to_format(value))
        end
        row.item("value#15").value(sum)  
      end
      # each user type
      5.downto(1) do |type|
        data_type = 322
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t("statistic_report.user_type_#{type}"))
          sum = 0
          hours.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND user_type = ? AND library_id = ? AND hour = ?", data_type, 0, type, 0, t+open])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value##{t+1}").value(to_format(value))
          end
          row.item("value#15").value(sum)  
          line_for_libraries(row) if type == 1
        end
      end
      # checkout users each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          hours.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 322 AND library_id = ? AND hour = ?", library.id, t+open]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value##{t+1}").value(to_format(value))
          end
          row.item("value#15").value(sum)
        end
        # each user type
        5.downto(1) do |type|
          sum = 0
          data_type = 322
          report.page.list(:list).add_row do |row|
            row.item(:option).value(I18n.t("statistic_report.user_type_#{type}"))
            hours.times do |t|
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND user_type = ? AND library_id = ? AND hour = ?", data_type, 0, type, library.id, t+open])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row.item("value##{t+1}").value(to_format(value))
            end
            row.item("value#15").value(sum)
            if type == 1
              if library == libraries.last
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end
        end
      end

      # checkout items all libraries
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.checkout_items'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        hours.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 321 AND library_id = ? AND hour = ?", 0, t+open]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value##{t+1}").value(to_format(value))
        end
        row.item("value#15").value(sum)  
      end
      3.times do |i|
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
          sum = 0
          hours.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND library_id = ? AND hour = ?", 321, i+1, 0, t+open])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value##{t+1}").value(to_format(value))
          end
          row.item("value#15").value(sum)  
          line_for_libraries(row) if i == 2
        end
      end
      # checkout items each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          hours.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 321 AND library_id = ? AND hour = ?", library.id, t+open]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value##{t+1}").value(to_format(value))
          end
          row.item("value#15").value(sum)
        end
        3.times do |i|
          report.page.list(:list).add_row do |row|
            row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
            sum = 0
            hours.times do |t|
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND library_id = ? AND hour = ?", 321, i+1, library.id, t+open])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row.item("value##{t+1}").value(to_format(value))
            end
            row.item("value#15").value(sum)
            if i == 2
              if library == libraries.last
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end
        end
      end

      # reserves all libraries
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.reserves'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        hours.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ?", 0, t+open]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value##{t+1}").value(to_format(value))
        end
        row.item("value#15").value(sum)  
      end
      # reserves on counter all libraries
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.on_counter'))
        sum = 0
        hours.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ? AND option = 1 AND age IS NULL", 0, t+open])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value##{t+1}").value(to_format(value))
        end
        row.item("value#15").value(sum)  
      end
      # reserves from OPAC all libraries
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.from_opac'))
        sum = 0
        hours.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ? AND option = 2 AND age IS NULL", 0, t+open])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value##{t+1}").value(to_format(value))
        end
        row.item("value#15").value(sum)  
        line_for_libraries(row)
      end
      # reserves each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          hours.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ?", library.id, t+open]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value##{t+1}").value(to_format(value))
          end
          row.item("value#15").value(sum)
        end
        # on counter
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.on_counter'))
          hours.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ? AND option = 1 AND age IS NULL", library.id, t+open])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value##{t+1}").value(to_format(value))
          end
          row.item("value#15").value(sum)
        end
        # from OPAC
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.from_opac'))
          hours.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ? AND option = 2 AND age IS NULL", library.id, t+open])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value##{t+1}").value(to_format(value))
          end
          row.item("value#15").value(sum)
          if library == libraries.last
            line(row)
          else
            line_for_libraries(row)
          end
        end
      end

      # questions all libraries
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.questions'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        hours.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 343 AND library_id = ? AND hour = ?", 0, t+open]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value##{t+1}").value(to_format(value))
        end
        row.item("value#15").value(sum)  
        row.item(:library_line).show
      end
      # reserves each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          hours.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 343 AND library_id = ? AND hour = ?", library.id, t+open]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value##{t+1}").value(to_format(value))
          end
          row.item("value#15").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_timezone_report_tsv(start_at, end_at)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{start_at}_#{end_at}_timezone_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:library, 'statistic_report.library'],
      [:option, 'statistic_report.option']
    ]
    #default setting 9 - 20
    open = SystemConfiguration.get("statistic_report.open")
    hours = SystemConfiguration.get("statistic_report.hours")

    libraries = Library.all
    logger.error "create daily timezone report: #{start_at} - #{end_at}"
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      hours.times do |t|
        row << "#{t+open}#{I18n.t('statistic_report.hour')}"
        columns << [t+open]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # checkout users all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkout_users')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 322 AND library_id = ? AND hour = ?", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # each user type
      5.downto(1) do |type|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << I18n.t("statistic_report.user_type_#{type}")
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND user_type = ? AND library_id = ? AND hour = ?", 322, 0, type, 0, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout users each libraries
      libraries.each do |library|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :library
            row << library.display_name
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 322 AND library_id = ? AND hour = ?", library.id, column[0]]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # each user type
        5.downto(1) do |type|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_users')
            when :library
              row << library.display_name
            when :option
              row << I18n.t("statistic_report.user_type_#{type}")
            when "sum"
              row << to_format(sum)
            else
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND user_type = ? AND library_id = ? AND hour = ?", 322, 0, type, library.id, column[0]])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # checkout items all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkout_items')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 321 AND library_id = ? AND hour = ?", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      3.times do |i|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << I18n.t("statistic_report.item_type_#{i+1}")
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND library_id = ? AND hour = ?", 321, i+1, 0, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout items each libraries
      libraries.each do |library|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :library
            row << library.display_name
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 321 AND library_id = ? AND hour = ?", library.id, column[0]]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        3.times do |i|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_items')
            when :library
              row << library.display_name
            when :option
              row << I18n.t("statistic_report.item_type_#{i+1}")
            when "sum"
              row << to_format(sum)
            else
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND library_id = ? AND hour = ?", 321, i+1, library.id, column[0]])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end

      # reserves all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ?", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves on counter all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << I18n.t('statistic_report.on_counter')
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ? AND option = 1 AND age IS NULL", 0, column[0]])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves from OPAC all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << I18n.t('statistic_report.from_opac')
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ? AND option = 2 AND age IS NULL", 0, column[0]])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves each libraries
      libraries.each do |library|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ?", library.id, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # on counter
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name
          when :option
            row << I18n.t('statistic_report.on_counter')
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ? AND option = 1 AND age IS NULL", library.id, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # from OPAC
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name
          when :option
            row << I18n.t('statistic_report.from_opac')
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 333 AND library_id = ? AND hour = ? AND option = 2 AND age IS NULL", library.id, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end

      # questions all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.questions')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 343 AND library_id = ? AND hour = ?", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves each libraries
      libraries.each do |library|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.questions')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 343 AND library_id = ? AND hour = ?", library.id, column[0]]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
    end
    return tsv_file
  end

  def self.get_day_report_pdf(start_at, end_at)
    libraries = Library.all
    logger.error "create day statistic report: #{start_at} - #{end_at}"

    begin
      report = ThinReports::Report.new :layout => get_layout_path("day_report")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      report.start_new_page
      report.page.item(:date).value(Time.now)
      report.page.item(:year_start_at).value(start_at[0,4])
      report.page.item(:month_start_at).value(start_at[4,2])
      report.page.item(:date_start_at).value(start_at[6,2])
      report.page.item(:year_end_at).value(end_at[0,4])
      report.page.item(:month_end_at).value(end_at[4,2])
      report.page.item(:date_end_at).value(end_at[6,2])

      # checkout users all libraries
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.checkout_users'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        7.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 222 AND library_id = ? AND day = ?", 0, t]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        row.item("valueall").value(sum)  
      end
      # each user type
      5.downto(1) do |type|
        data_type = 222
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t("statistic_report.user_type_#{type}"))
          sum = 0
          7.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND user_type = ? AND library_id = ? AND day = ?", data_type, 0, type, 0, t])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          row.item("valueall").value(sum)  
          line_for_libraries(row) if type == 1
        end
      end

      # checkout users each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          7.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 222 AND library_id = ? AND day = ?", library.id, t]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          row.item("valueall").value(sum)
        end
        # each user type
        5.downto(1) do |type|
          sum = 0
          data_type = 222
          report.page.list(:list).add_row do |row|
            row.item(:option).value(I18n.t("statistic_report.user_type_#{type}"))
            7.times do |t|
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND user_type = ? AND library_id = ? AND day = ?", data_type, 0, type, library.id, t])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row.item("value#{t}").value(to_format(value))
            end
            row.item("valueall").value(sum)
            if type == 1
              if library == libraries.last
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end
        end
      end

      # checkout items all libraries
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.checkout_items'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        7.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 221 AND library_id = ? AND day = ?", 0, t]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        row.item("valueall").value(sum)  
      end
      3.times do |i|
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
          sum = 0
          7.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND library_id = ? AND day = ?", 221, i+1, 0, t])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          row.item("valueall").value(sum)  
          line_for_libraries(row) if i == 2
        end
      end
      # checkout items each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          7.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 221 AND library_id = ? AND day = ?", library.id, t]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          row.item("valueall").value(sum)
        end
        3.times do |i|
          report.page.list(:list).add_row do |row|
            row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
            sum = 0
            7.times do |t|
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND library_id = ? AND day = ?", 221, i+1, library.id, t])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row.item("value#{t}").value(to_format(value))
            end
            row.item("valueall").value(sum)
            if i == 2
              if library == libraries.last
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end
        end
      end

      # reserves all libraries
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.reserves'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        7.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ?", 0, t]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        row.item("valueall").value(sum)  
      end
      # reserves on counter all libraries
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.on_counter'))
        sum = 0
        7.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ? AND option = 1", 0, t]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        row.item("valueall").value(sum)  
      end
      # reserves from OPAC all libraries
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.from_opac'))
        sum = 0
        7.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ? AND option = 2", 0, t]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        row.item("valueall").value(sum)  
        line_for_libraries(row)
      end
      # reserves each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name.localize)
          7.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ?", library.id, t])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          row.item("valueall").value(sum)
        end
        # on counter
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.on_counter'))
          7.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ? AND option = 1", library.id, t])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          row.item("valueall").value(sum)
        end
        # from OPAC
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.from_opac'))
          7.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ? AND option = 2", library.id, t])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          row.item("valueall").value(sum)
          if library == libraries.last
            line(row)
          else
            line_for_libraries(row) 
          end
        end
      end
     
 
      # questions all libraries
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.questions'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        7.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 243 AND library_id = ? AND day = ?", 0, t]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        row.item("valueall").value(sum)  
        row.item(:library_line).show
      end
      # questions each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          7.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 243 AND library_id = ? AND day = ?", library.id, t]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_day_report_tsv(start_at, end_at)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{start_at}_#{end_at}_day_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:library, 'statistic_report.library'],
      [:option, 'statistic_report.option']
    ]
    libraries = Library.all
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      7.times do |t|
        row << I18n.t("statistic_report.day_#{t}")
        columns << [t]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # checkout users all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkout_users')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 222 AND library_id = ? AND day = ?", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # each user type
      5.downto(1) do |type|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << I18n.t("statistic_report.user_type_#{type}")
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND user_type = ? AND library_id = ? AND day = ?", 222, 0, type, 0, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout users each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :library
            row << library.display_name.localize
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 222 AND library_id = ? AND day = ?", library.id, column[0]]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # each user type
        5.downto(1) do |type|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_users')
            when :library
              row << library.display_name.localize
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND user_type = ? AND library_id = ? AND day = ?", 222, 0, type, library.id, column[0]])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # checkout items all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkout_items')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 221 AND library_id = ? AND day = ?", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      3.times do |i|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << I18n.t("statistic_report.item_type_#{i+1}")
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND library_id = ? AND day = ?", 221, i+1, 0, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout items each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :library
            row << library.display_name.localize
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 221 AND library_id = ? AND day = ?", library.id, column[0]]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        3.times do |i|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_items')
            when :library
              row << library.display_name.localize
            when :option
              row << I18n.t("statistic_report.item_type_#{i+1}")
            when "sum"
              row << to_format(sum)
            else
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND library_id = ? AND day = ?", 221, i+1, library.id, column[0]])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # reserves all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ?", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves on counter all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << I18n.t('statistic_report.on_counter')
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ? AND option = 1", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves from OPAC all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << I18n.t('statistic_report.from_opac')
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ? AND option = 2", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name.localize
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ?", library.id, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # on counter
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name.localize
          when :option
            row << I18n.t('statistic_report.on_counter')
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ? AND option = 1", library.id, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # from OPAC
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name.localize
          when :option
            row << I18n.t('statistic_report.from_opac')
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 233 AND library_id = ? AND day = ? AND option = 2", library.id, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end 
      # questions all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.questions')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 243 AND library_id = ? AND day = ?", 0, column[0]]).no_condition
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # questions each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.questions')
          when :library
            row << library.display_name.localize
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = 243 AND library_id = ? AND day = ?", library.id, column[0]]).no_condition
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
    end
    return tsv_file
  end

  def self.get_age_report_pdf(start_at, end_at)
    libraries = Library.all
    logger.error "create day statistic report: #{start_at} - #{end_at}"

    begin
      report = ThinReports::Report.new :layout => get_layout_path("age_report")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      report.start_new_page
      report.page.item(:date).value(Time.now)
      report.page.item(:year_start_at).value(start_at[0,4])
      report.page.item(:month_start_at).value(start_at[4,2])
      report.page.item(:date_start_at).value(start_at[6,2])
      report.page.item(:year_end_at).value(end_at[0,4])
      report.page.item(:month_end_at).value(end_at[4,2])
      report.page.item(:date_end_at).value(end_at[6,2])

      # checkout users all libraries
      data_type = 222
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.checkout_users'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ?", data_type, t, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ?", data_type, 10, 0])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))  
        row.item("valueall").value(sum)  
        row.item(:library_line).show
      end
      # checkout users each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ?", data_type, t, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ?", data_type, 10, library.id])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end

      # checkout items all libraries
      data_type = 221
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.checkout_items'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = 0", data_type, t, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = 0", data_type, 10, 0])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))  
        row.item("valueall").value(sum)  
      end
      3.times do |i|
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
          sum = 0
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = ?", data_type, t, 0, i+1])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = ?", data_type, 10, 0, i+1])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))  
          row.item("valueall").value(sum)  
          line_for_libraries(row) if i == 2
        end
      end
      # checkout items each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = 0", data_type, t, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = 0", data_type, 10, library.id])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
        end
        3.times do |i|
          report.page.list(:list).add_row do |row|
            row.item(:option).value(I18n.t("statistic_report.item_type_#{i+1}"))
            sum = 0
            8.times do |t|
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = ?", data_type, t, library.id, i+1])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row.item("value#{t}").value(to_format(value))
            end
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = ?", data_type, 10, library.id, i+1])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value8").value(to_format(value))
            row.item("valueall").value(sum)
            if i == 2
              if library == libraries.last
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end  
        end
      end

      # all users all libraries
      data_type = 212
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.users'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        #row.item(:option).value(I18n.t('statistic_report.all_users'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 0, t, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end  
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 0, 10, 0])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))
        row.item("valueall").value(sum)
      end
      # unlocked users all libraries
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.unlocked_users'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 1, t, 0])
          datas.each do |data|
            value =value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end  
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 1, 10, 0])
        datas.each do |data|
          value =value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))
        row.item("valueall").value(sum)
      end
      # locked users all libraries
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.locked_users'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 2, t, 0])
          datas.each do |data|
            value =value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end  
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 2, 10, 0])
        datas.each do |data|
          value =value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))
        row.item("valueall").value(sum)
      end
      # provisional users all libraries
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.user_provisional'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 3, t, 0])
          datas.each do |data|
            value =value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end  
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 3, 10, 0])
        datas.each do |data|
          value =value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))
        row.item("valueall").value(sum)
        line_for_libraries(row)
      end
      # users each libraries
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name.localize)
          sum = 0
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 0, t, library.id])
            datas.each do |data|
              value =value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end  
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 0, 10, library.id])
          datas.each do |data|
            value =value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
        end
        # unlocked users each libraries
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.unlocked_users'))
          sum = 0
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 1, t, library.id])
            datas.each do |data|
              value =value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end    
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 1, 10, library.id])
          datas.each do |data|
            value =value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
        end
        # locked users each libraries
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.locked_users'))
          sum = 0
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 2, t, library.id])
            datas.each do |data|
              value =value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end  
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 2, 10, library.id])
          datas.each do |data|
            value =value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
        end
        # provisional users each libraries
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.user_provisional'))
          sum = 0
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 3, t, library.id])
            datas.each do |data|
              value =value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end  
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", data_type, 3, 10, library.id])
          datas.each do |data|
            value =value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
          if library == libraries.last
            line(row)
          else
            line_for_libraries(row)
          end
        end
      end

      # user_areas all libraries
      data_type = 212
      # all_area
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.user_areas'))
        row.item(:library).value(I18n.t('statistic_report.all_areas'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND area_id = ?", data_type, t, 0])
          #datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND age = ?", end_at, data_type, t])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND area_id = ?", data_type, 10, 0])
        #datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND age = ?", end_at, data_type, 10])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))
        row.item("valueall").value(sum)
        row.item(:library_line).show
      end
      # each_area
      @areas = Area.all
      @areas.each do |a|
        report.page.list(:list).add_row do |row|
          row.item(:library).value(a.name)
          sum = 0
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND area_id = ? AND age = ?", data_type, a.id, t])
            #datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND area_id = ? AND age = ?", end_at, data_type, a.id, t])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND area_id = ? AND age = ?", data_type, a.id, 10])
          #datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND area_id = ? AND age = ?", end_at, data_type, a.id, 10])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))

          row.item("valueall").value(sum)
          row.item(:library_line).show
        end
      end
      # unknown area
      report.page.list(:list).add_row do |row|
        row.item(:library).value(I18n.t('statistic_report.other_area'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND area_id = ? AND option = ?", data_type, t, 0, 3])
          #datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND age = ? AND area_id = 0", end_at, data_type, t])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND area_id = ? AND option = ?", data_type, 10, 0, 3])
        #datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND age = ? AND area_id = 0", end_at, data_type, 10])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))
        row.item("valueall").value(sum)
        line(row)
      end

      # reserves all libraries
      data_type = 233
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.reserves'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", data_type, t, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", data_type, 10, 0])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))  
        row.item("valueall").value(sum)  
      end
      # reserves on counter all libraries
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.on_counter'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 1 AND age = ? AND library_id = ?", data_type, t, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 1 AND age = ? AND library_id = ?", data_type, 10, 0])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))  
        row.item("valueall").value(sum)  
      end
      # reserves from OPAC all libraris
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('statistic_report.from_opac'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 2 AND age = ? AND library_id = ?", data_type, t, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 2 AND age = ? AND library_id = ?", data_type, 10, 0])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))  
        row.item("valueall").value(sum)  
        line_for_libraries(row)
      end
      # reserves each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", data_type, t, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", data_type, 10, library.id])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
        end
        # on counter
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.on_counter'))
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 1 AND age = ? AND library_id = ?", data_type, t, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 1 AND age = ? AND library_id = ?", data_type, 10, library.id])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
        end
        # from OPAC
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('statistic_report.from_opac'))
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 2 AND age = ? AND library_id = ?", data_type, t, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 2 AND age = ? AND library_id = ?", data_type, 10, library.id])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
          if library == libraries.last
            line(row)
          else
            line_for_libraries(row)
          end
        end
      end

      # questions all libraries
      data_type = 243
      report.page.list(:list).add_row do |row|
        row.item(:type).value(I18n.t('statistic_report.questions'))
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        8.times do |t|
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", data_type, t, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value#{t}").value(to_format(value))
        end
        value = 0
        datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", data_type, 10, 0])
        datas.each do |data|
          value = value + data.value
        end
        sum = sum + value
        row.item("value8").value(to_format(value))  
        row.item("valueall").value(sum)  
        row.item(:library_line).show
      end
      # questions each libraries
      libraries.each do |library|
        sum = 0
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          8.times do |t|
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", data_type, t, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row.item("value#{t}").value(to_format(value))
          end
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", data_type, 10, library.id])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row.item("value8").value(to_format(value))
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_age_report_tsv(start_at, end_at)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{start_at}_#{end_at}_day_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:library, 'statistic_report.library'],
      [:area, 'activerecord.models.area'],
      [:option, 'statistic_report.option']
    ]
    libraries = Library.all
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      8.times do |t|
        if t == 0
          row << "#{t} ~ 9"
        elsif t < 7
          row << "#{t}0 ~ #{t}9"
        else
          row << "#{t}0 ~ "
        end
        columns << [t]
      end
      row << I18n.t('page.unknown')
      columns << ["unknown"]
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # checkout users all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkout_users')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << ""
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ?", 222, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ?", 222, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # checkout users each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :library
            row << library.display_name
          when :area
            row << ""
          when :option
            row << ""
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ?", 222, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ?", 222, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout items all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkout_items')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << ""
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = 0", 221, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = 0", 221, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      3.times do |i|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :area
            row << ""
          when :option
            row << I18n.t("statistic_report.item_type_#{i+1}")
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = ?", 221, 10, 0, i+1])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = ?", 221, column[0], 0, i+1])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout items each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :library
            row << library.display_name
          when :area
            row << ""
          when :option
            row << ""
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = 0", 221, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = 0", 221, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        3.times do |i|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_items')
            when :library
              row << library.display_name
            when :area
              row << ""
            when :option
              row << I18n.t("statistic_report.item_type_#{i+1}")
            when "unknown"
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = ?", 221, 10, library.id, i+1])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row << to_format(value)
            when "sum"
              row << to_format(sum)
            else
              value = 0
              datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND age = ? AND library_id = ? AND option = ?", 221, column[0], library.id, i+1])
              datas.each do |data|
                value = value + data.value
              end
              sum = sum + value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # all users all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.users')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << I18n.t('statistic_report.all_users')
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 0, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 0, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # unlocked users all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.users')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << I18n.t('statistic_report.unlocked_users')
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 1, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 1, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # locked users all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.users')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << I18n.t('statistic_report.locked_users')
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 2, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 2, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # provisional users all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.users')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << I18n.t('statistic_report.user_provisional')
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 3, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 3, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # users each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.users')
          when :library
            row << library.display_name.localize
          when :area
            row << ""
          when :option
            row << ""
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 0, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 0, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # unlocked users each libraries
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.users')
          when :library
            row << library.display_name.localize
          when :area
            row << ""
          when :option
            row << I18n.t('statistic_report.unlocked_users')
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 1, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 1, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # locked users each libraries
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.users')
          when :library
            row << library.display_name.localize
          when :area
            row << ""
          when :option
            row << I18n.t('statistic_report.locked_users')
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 2, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 2, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # provisional users each libraries
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.users')
          when :library
            row << library.display_name.localize
          when :area
            row << ""
          when :option
            row << I18n.t('statistic_report.user_provisional')
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 3, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = ? AND age = ? AND library_id = ?", 212, 3, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # user_areas all libraries
      # all_area
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.user_areas')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << I18n.t('statistic_report.all_areas')
        when :option
          row << ""
        when "unknown"
          value = 0
        datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND age = ?", end_at, 262, 10])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND age = ?", end_at, 262, column[0]])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # each_area
      @areas = Area.all
      @areas.each do |a|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.user_areas')
          when :library
            row << ""
          when :area
            row << a.name
          when :option
            row << ""
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND area_id = ? AND age = ?", end_at, 262, a.id, 10])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND area_id = ? AND age = ?", end_at, 262, a.id, column[0]])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # unknown area
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.user_areas')
        when :library
          row << ""
        when :area
          row << I18n.t('statistic_report.other_area')
        when :option
          row << ""
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND age = ? AND area_id = 0", end_at, 262, 10])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd = ? AND data_type = ? AND age = ? AND area_id = 0", end_at, 262, column[0]])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << ""
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", 233, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", 233, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves on counter all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << I18n.t('statistic_report.on_counter')
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 1 AND age = ? AND library_id = ?", 233, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 1 AND age = ? AND library_id = ?", 233, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves from OPAC all libraris
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << I18n.t('statistic_report.from_opac')
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 2 AND age = ? AND library_id = ?", 233, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 2 AND age = ? AND library_id = ?", 233, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # reserves each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name.localize
          when :area
            row << ""
          when :option
            row << ""
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", 233, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", 233, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # on counter
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name.localize
          when :area
            row << ""
          when :option
            row << I18n.t('statistic_report.on_counter')
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 1 AND age = ? AND library_id = ?", 233, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 1 AND age = ? AND library_id = ?", 233, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # from OPAC
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')
          when :library
            row << library.display_name.localize
          when :area
            row << ""
          when :option
            row << I18n.t('statistic_report.from_opac')
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 2 AND age = ? AND library_id = ?", 233, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 2 AND age = ? AND library_id = ?", 233, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # questions all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.questions')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :area
          row << ""
        when :option
          row << ""
        when "unknown"
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", 243, 10, 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        when "sum"
          row << to_format(sum)
        else
          value = 0
          datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", 243, column[0], 0])
          datas.each do |data|
            value = value + data.value
          end
          sum = sum + value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # questions each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.questions')
          when :library
            row << library.display_name.localize
          when :area
            row << ""
          when :option
            row << ""
          when "unknown"
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", 243, 10, library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          when "sum"
            row << to_format(sum)
          else
            value = 0
            datas = Statistic.where(["yyyymmdd >= #{start_at} AND yyyymmdd <= #{end_at} AND data_type = ? AND option = 0 AND age = ? AND library_id = ?", 243, column[0], library.id])
            datas.each do |data|
              value = value + data.value
            end
            sum = sum + value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
    end
    return tsv_file
  end

  def self.get_items_daily_pdf(term)
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    begin 
      report = ThinReports::Report.new :layout => get_layout_path("items_daily")

      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      num_for_last_page = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i - 26
      [1,14,27].each do |start_date| # for 3 pages
        report.start_new_page
        report.page.item(:date).value(Time.now)
        report.page.item(:year).value(term[0,4])
        report.page.item(:month).value(term[4,6])        
        # header
        if start_date != 27
          13.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
        else
          num_for_last_page.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => start_date))
          end
          report.page.list(:list).header.item("column#13").value(I18n.t('statistic_report.sum'))
        end
        # items all libraries
        data_type = 211
        report.page.list(:list).add_row do |row|
          row.item(:library).value(I18n.t('statistic_report.all_library'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
              row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
            end
          end
          row.item("condition_line").show
        end  
        # items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
              row.item(:option).value(num)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :call_number => num).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :call_number => num).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
                end
              end
              row.item("condition_line").show if num == call_numbers.last
            end  
          end
        end
        # items each checkout_types
        checkout_types.each do |checkout_type|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
            row.item(:option).value(checkout_type.display_name.localize)
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
              end
            end
            row.item("condition_line").show if checkout_type == checkout_types.last
          end  
        end
        # missing items
        report.page.list(:list).add_row do |row|
          row.item(:condition).value(I18n.t('statistic_report.missing_items'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :option => 1, :library_id => 0).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :option => 1, :library_id => 0).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
              row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
            end
          end
          line_for_items(row)
        end  
        # items each library
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:library).value(library.display_name)
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
                row.item("value##{t+1}").value(to_format(value))
              end 
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
              end
            end
            row.item("condition_line").show
          end  
          # items each call_numbers
          unless call_numbers.nil?
            call_numbers.each do |num|
              report.page.list(:list).add_row do |row|
                row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                row.item(:option).value(num)
                if start_date != 27
                  13.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :call_number => num).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                  end
                else
                  num_for_last_page.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :call_number => num).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                    row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
                  end
                end
                row.item("condition_line").show if num == call_numbers.last
              end
            end
          end
          # items each checkout_types
          checkout_types.each do |checkout_type|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
              row.item(:option).value(checkout_type.display_name.localize)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
                end
              end
              row.item(:condition_line).show if checkout_type == checkout_types.last
            end
          end
          # missing items
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('statistic_report.missing_items'))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymm => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :option => 1, :library_id => library.id).first.value rescue 0 
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :option => 1, :library_id => library.id).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
              end
            end
            row.item(:library_line).show
            row.item(:condition_line).show
            line_for_items(row) if library.shelves.size < 1
          end  
          # items each shelves and call_numbers
          library.shelves.each do |shelf|
            report.page.list(:list).add_row do |row|
              row.item(:library).value("(#{shelf.display_name})")
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
                end
              end
              row.item("library_line").show
              row.item("condition_line").show
              line_for_items(row) if call_numbers.nil? and shelf == library.shelves.last
            end
            unless call_numbers.nil?
              call_numbers.each do |num|
                report.page.list(:list).add_row do |row|
                  row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                  row.item(:option).value(num)
                  if start_date != 27
                    13.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                    end
                  else
                    num_for_last_page.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                      row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
                    end
                  end
                  if num == call_numbers.last
                    row.item("library_line").show
                    row.item("condition_line").show
                    line_for_items(row) if shelf == library.shelves.last
                  end
                end
              end
            end
          end
        end
      end

      return report.generate
      return true
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_items_daily_tsv(term)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_items_daily_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:library, 'statistic_report.library'],
      [:shelf, 'activerecord.models.shelf'],
      [:condition, 'statistic_report.condition'],
      [:option, 'statistic_report.option'] 
    ]
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    days = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      days.times do |t|
        row << I18n.t('statistic_report.date', :num => t+1)
        columns << ["#{term}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      # items all libraries
      row = []
      columns.each do |column|
        case column[0]
        when :library 
          row << I18n.t('statistic_report.all_library')
        when :shelf
          row << ""
        when :condition
          row << ""
        when :option
          row << ""
        when "sum"
          logger.error "sum: #{term}#{days+1}"
          value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
          row << to_format(value)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # items each call_numbers
      unless call_numbers.nil?
        call_numbers.each do |num|
          row = []
          columns.each do |column|
            case column[0]
            when :library 
              row << I18n.t('statistic_report.all_library')
            when :shelf
              row << ""
            when :condition
              row << I18n.t('activerecord.attributes.item.call_number')
            when :option
              row << num
            when "sum"
              value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :library_id => 0, :call_number => num).first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => 0, :call_number => num).first.value rescue 0
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # items each checkout_types
      checkout_types.each do |checkout_type|
        row = []
        columns.each do |column|
          case column[0]
          when :library 
            row << I18n.t('statistic_report.all_library')
          when :shelf
            row << ""
          when :condition
            row << I18n.t('activerecord.models.checkout_type')
          when :option
            row << checkout_type.display_name.localize
          when "sum"
            value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
            row << to_format(value)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # missing items
      row = []
      columns.each do |column|
        case column[0]
        when :library 
          row << I18n.t('statistic_report.all_library')
        when :shelf
          row << ""
        when :condition
          row << I18n.t('statistic_report.missing_items')
        when :option
          row << ""
        when "sum"
          value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :option => 1, :library_id => 0).first.value rescue 0
          row << to_format(value)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :option => 1, :library_id => 0).first.value rescue 0
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # items each library
      libraries.each do |library|
        row = []
        columns.each do |column|
          case column[0]
          when :library 
            row << library.display_name
          when :shelf
            row << ""
          when :condition
            row << ""
          when :option
            row << ""
          when "sum"
            value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :library_id => library.id).no_condition.first.value rescue 0 
            row << to_format(value)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => library.id).no_condition.first.value rescue 0 
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            row = []
            columns.each do |column|
              case column[0]
              when :library 
                row << library.display_name
              when :shelf
                row << ""
              when :condition
                row << I18n.t('activerecord.attributes.item.call_number')
              when :option
                row << num
              when "sum"
                value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :library_id => library.id, :call_number => num).first.value rescue 0
                row << to_format(value)
              else
                value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => library.id, :call_number => num).first.value rescue 0
                row << to_format(value)
              end
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
        # items each checkout_types
        checkout_types.each do |checkout_type|
          row = []
          columns.each do |column|
            case column[0]
            when :library 
              row << library.display_name
            when :shelf
              row << ""
            when :condition
              row << I18n.t('activerecord.models.checkout_type')
            when :option
              row << checkout_type.display_name.localize
            when "sum"
              value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # missing items
        row = []
        columns.each do |column|
          case column[0]
          when :library 
            row << library.display_name
          when :shelf
            row << ""
          when :condition
            row << I18n.t('statistic_report.missing_items')
          when :option
            row << ""
          when "sum"
            value = Statistic.where(:yyyymm => "#{term}#{days+1}", :data_type => 211, :option => 1, :library_id => library.id).first.value rescue 0 
            row << to_format(value)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 211, :option => 1, :library_id => library.id).first.value rescue 0 
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # items each shelves and call_numbers
        library.shelves.each do |shelf|
          row = []
          columns.each do |column|
            case column[0]
            when :library 
              row << library.display_name.localize
            when :shelf
              row << shelf.display_name.localize
            when :condition
              row << I18n.t('activerecord.models.checkout_type')
            when :option
              row << ""
            when "sum"
              value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :library_id => library.id, :shelf_id => shelf.id).first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => library.id, :shelf_id => shelf.id).first.value rescue 0
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless call_numbers.nil?
            call_numbers.each do |num|
              row = []
              columns.each do |column|
                case column[0]
                when :library 
                  row << library.display_name.localize
                when :shelf
                  row << shelf.display_name.localize
                when :condition
                  row << I18n.t('activerecord.attributes.item.call_number')
                when :option
                  row << num
                when "sum"
                  value = Statistic.where(:yyyymmdd => "#{term}#{days+1}", :data_type => 211, :library_id => library.id, :shelf_id => shelf.id, :call_number => num).first.value rescue 0
                  row << to_format(value)
                else
                  value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => library.id, :shelf_id => shelf.id, :call_number => num).first.value rescue 0
                  row << to_format(value)
                end
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
        end
      end
    end
    return tsv_file
  end

  def self.get_items_monthly_pdf(term)
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    begin 
      report = ThinReports::Report.new :layout => get_layout_path("items_monthly")

      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      report.start_new_page
      report.page.item(:date).value(Time.now)       
      report.page.item(:term).value(term)

      # items all libraries
      data_type = 111
      report.page.list(:list).add_row do |row|
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        12.times do |t|
          if t < 3 # for Japanese fiscal year
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
          else
            value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
          end
          row.item("value#{t+1}").value(to_format(value))
          row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
          row.item("condition_line").show
        end  
      end
      # items each call_numbers
      unless call_numbers.nil?
        call_numbers.each do |num|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
            row.item(:option).value(num)
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :call_number => num).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :call_number => num).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              row.item("condition_line").show if num == call_numbers.last
            end  
          end
        end
      end
      # items each checkout_types
      checkout_types.each do |checkout_type|
        report.page.list(:list).add_row do |row|
          row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
          row.item(:option).value(checkout_type.display_name.localize)
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            row.item("condition_line").show if checkout_type == checkout_types.last
          end  
        end
      end
      # missing items
      report.page.list(:list).add_row do |row|
        row.item(:condition).value(I18n.t('statistic_report.missing_items'))
        12.times do |t|
          if t < 3 # for Japanese fiscal year
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :option => 1, :library_id => 0).first.value rescue 0
          else
            value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :option => 1, :library_id => 0).first.value rescue 0
          end
          row.item("value#{t+1}").value(to_format(value))
          row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
          line_for_items(row)
        end  
      end
      # items each library
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            row.item("condition_line").show
          end  
        end
        # items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
              row.item(:option).value(num)
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  datas = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :call_number => num)
                else
                  datas = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :call_number => num)
                end
                value = 0
                datas.each do |data|
                  value += data.value
                end
                row.item("value#{t+1}").value(to_format(value))
                row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
                row.item("condition_line").show if num == call_numbers.last
              end
            end
          end
        end
        # items each checkout_types
        checkout_types.each do |checkout_type|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
            row.item(:option).value(checkout_type.display_name.localize)
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              row.item(:condition_line).show if checkout_type == checkout_types.last
            end
          end
        end
        # missing items
        report.page.list(:list).add_row do |row|
          row.item(:condition).value(I18n.t('statistic_report.missing_items'))
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :option => 1, :library_id => library.id).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :option => 1, :library_id => library.id).first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            row.item(:library_line).show
            row.item(:condition_line).show
            line_for_items(row) if library.shelves.size < 1
          end  
        end
        # items each shelves and call_numbers
        library.shelves.each do |shelf|
          report.page.list(:list).add_row do |row|
            row.item(:library).value("(#{shelf.display_name})")
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                datas = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id)
              else
                datas = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id)
              end
              value = 0
              datas.each do |data|
                value += data.value
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
              row.item("library_line").show
              row.item("condition_line").show
            end
            line_for_items(row) if shelf == library.shelves.last && call_numbers.nil?
          end
          unless call_numbers.nil?
            call_numbers.each do |num|
              report.page.list(:list).add_row do |row|
                row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                row.item(:option).value(num)
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
                  if num == call_numbers.last
                    row.item("library_line").show
                    row.item("condition_line").show
                    line_for_items(row) if shelf == library.shelves.last
                  end
                end
              end
            end
          end
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_items_monthly_tsv(term)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_items_monthly_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:library, 'statistic_report.library'],
      [:shelf, 'activerecord.models.shelf'],
      [:condition, 'statistic_report.condition'],
      [:option, 'statistic_report.option'] 
    ]
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      9.times do |t|
        row << I18n.t('statistic_report.month', :num => t+4)
        columns << ["#{term}#{"%02d" % (t + 4)}"]
      end
      3.times do |t|
        row << I18n.t('statistic_report.month', :num => t+1)
        columns << ["#{term.to_i + 1}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # items all libraries
      row = []
      columns.each do |column|
        case column[0]
        when :library 
          row << I18n.t('statistic_report.all_library')
        when :shelf
          row << ""
        when :condition
          row << ""
        when :option
          row << ""
        when "sum"
          value = Statistic.where(:yyyymm => "#{term.to_i+1}03}", :data_type => 111, :library_id => 0).no_condition.first.value rescue 0
          row << to_format(value)
        else
          value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => 0).no_condition.first.value rescue 0
          row << to_format(value)
        end   
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # items each call_numbers
      unless call_numbers.nil?
        call_numbers.each do |num|
          row = []
          columns.each do |column|
            case column[0]
            when :library 
              row << I18n.t('statistic_report.all_library')
            when :shelf
              row << ""
            when :condition
              row << I18n.t('activerecord.attributes.item.call_number')
            when :option
              row << num
            when "sum"
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 111, :library_id => 0, :call_number => num).first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :call_number => num).first.value rescue 0
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end   
      end
      # items each checkout_types
      checkout_types.each do |checkout_type|
        row = []
        columns.each do |column|
          case column[0]
          when :library 
            row << I18n.t('statistic_report.all_library')
          when :shelf
            row << ""
          when :condition
            row << I18n.t('activerecord.models.checkout_type')
          when :option
            row << checkout_type.display_name.localize
          when "sum"
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}03", :data_type => 111, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
            row << to_format(value)
          else
              value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => 0, :checkout_type_id => checkout_type.id).first.value rescue 0
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # missing items
      row = []
      columns.each do |column|
        case column[0]
        when :library 
          row << I18n.t('statistic_report.all_library')
        when :shelf
          row << ""
        when :condition
          row << I18n.t('statistic_report.missing_items')
        when :option
          row << ""
        when "sum"
          value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 111, :option => 1, :library_id => 0).first.value rescue 0
          row << to_format(value)
        else
            value = Statistic.where(:yyyymm => column[0], :data_type => 111, :option => 1, :library_id => 0).first.value rescue 0
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # items each library
      libraries.each do |library|
        row = []
        columns.each do |column|
          case column[0]
          when :library 
            row << library.display_name
          when :shelf
            row << ""
          when :condition
            row << ""
          when :option
            row << ""
          when "sum"
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}03", :data_type => 111, :library_id => library.id).no_condition.first.value rescue 0 
            row << to_format(value)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id).no_condition.first.value rescue 0 
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            row = []
            columns.each do |column|
              case column[0]
              when :library 
                row << library.display_name
              when :shelf
                row << ""
              when :condition
                row << I18n.t('activerecord.attributes.item.call_number')
              when :option
                row << num
              when "sum"
                datas = Statistic.where(:yyyymm => "#{term.to_i + 1}03", :data_type => 111, :library_id => library.id, :call_number => num)
                value = 0
                datas.each do |data|
                  value += data.value
                end
                row << to_format(value)
              else
                datas = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :call_number => num)
                value = 0
                datas.each do |data|
                  value += data.value
                end
                row << to_format(value)
              end
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
        # items each checkout_types
        checkout_types.each do |checkout_type|
          row = []
          columns.each do |column|
            case column[0]
            when :library 
              row << library.display_name
            when :shelf
              row << ""
            when :condition
              row << I18n.t('activerecord.models.checkout_type')
            when :option
              row << checkout_type.display_name.localize
            when "sum"
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}03", :data_type => 111, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
              row << to_format(value)
            else
                value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :checkout_type_id => checkout_type.id).first.value rescue 0
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # missing items
        row = []
        columns.each do |column|
          case column[0]
          when :library 
            row << library.display_name
          when :shelf
            row << ""
          when :condition
            row << I18n.t('statistic_report.missing_items')
          when :option
            row << ""
          when "sum"
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}03", :data_type => 111, :option => 1, :library_id => library.id).first.value rescue 0 
            row << to_format(value)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 111, :option => 1, :library_id => library.id).first.value rescue 0 
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # items each shelves and call_numbers
        library.shelves.each do |shelf|
          row = []
          columns.each do |column|
            case column[0]
            when :library 
              row << library.display_name
            when :shelf
              row << shelf.display_name.localize
            when :condition
              row << ""
            when :option
              row << ""
            when "sum"
              datas = Statistic.where(:yyyymm => "#{term.to_i + 1}03", :data_type => 111, :library_id => library.id, :shelf_id => shelf.id)
              value = 0
              datas.each do |data|
                value += data.value
              end
              row << to_format(value)
            else
              datas = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :shelf_id => shelf.id)
              value = 0
              datas.each do |data|
                value += data.value
              end
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless call_numbers.nil?
            call_numbers.each do |num|
              row = []
              columns.each do |column|
                case column[0]
                when :library 
                  row << library.display_name
                when :shelf
                  row << shelf.display_name.localize
                when :condition
                  row << I18n.t('activerecord.attributes.item.call_number')
                when :option
                  row << num
                when "sum"
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}03", :data_type => 111, :library_id => library.id, :shelf_id => shelf.id, :call_number => num).first.value rescue 0
                  row << to_format(value)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :shelf_id => shelf.id, :call_number => num).first.value rescue 0
                  row << to_format(value)
                end
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
        end
      end
    end
    return tsv_file
  end

  def self.get_inout_daily_pdf(term)
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    logger.error "create daily inout items statistic report: #{term}"

    begin
      report = ThinReports::Report.new :layout => get_layout_path("inout_items_daily")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      num_for_last_page = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i - 26
      [1,14,27].each do |start_date| # for 3 pages
        # accept items
        report.start_new_page
        report.page.item(:date).value(Time.now)
        report.page.item(:year).value(term[0,4])
        report.page.item(:month).value(term[4,6])        
        report.page.item(:inout_type).value(I18n.t('statistic_report.accept'))
        # header
        if start_date != 27
          13.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
        else
          num_for_last_page.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
          report.page.list(:list).header.item("column#13").value(I18n.t('statistic_report.sum'))
        end
        # accept items all libraries
        data_type = 211
        report.page.list(:list).add_row do |row|
          row.item(:library).value(I18n.t('statistic_report.all_library'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 2).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 2).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
              if t == num_for_last_page - 1
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :option => 2)
                datas.each do |data|
                  sum += data.value
                end
                row.item("value#13").value(sum)
              end
            end
          end
          row.item(:condition_line).show
        end
        # accept items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
              row.item(:option).value(num)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :call_number => num, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :call_number => num, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  if t == num_for_last_page - 1
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :call_number => num, :option => 2)
                    datas.each do |data|
                      sum += data.value
                    end
                    row.item("value#13").value(sum)
                  end
                end
              end
              row.item("condition_line").show if num == call_numbers.last
            end  
          end
        end
        # accept items each checkout_types
        checkout_types.each do |checkout_type|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
            row.item(:option).value(checkout_type.display_name.localize)
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                if t == num_for_last_page - 1
                  sum = 0
                  datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 2)
                  datas.each do |data|
                    sum += data.value
                  end
                  row.item("value#13").value(sum)
                end
              end
            end  
            row.item("condition_line").show if checkout_type == checkout_types.last
            line_for_items(row) if checkout_type == checkout_types.last
          end
        end
        # accept items each libraries
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:library).value(library.display_name)
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 2).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 2).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                if t == num_for_last_page - 1
                  sum = 0
                  datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :option => 2)
                  datas.each do |data|
                    sum += data.value
                  end
                  row.item("value#13").value(sum)
                end
              end
            end
            row.item(:condition_line).show
          end
          # accept items each call_numbers
          unless call_numbers.nil?
            call_numbers.each do |num|
              report.page.list(:list).add_row do |row|
                row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                row.item(:option).value(num)
                if start_date != 27
                  13.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :call_number => num, :option => 2).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                  end
                else
                  num_for_last_page.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :call_number => num, :option => 2).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                    if t == num_for_last_page - 1
                      sum = 0
                      datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :call_number => num, :option => 2)
                      datas.each do |data|
                        sum += data.value
                      end
                      row.item("value#13").value(sum)
                    end
                  end
                end
                row.item("condition_line").show if num == call_numbers.last
              end
            end
          end
          # accept items each checkout_types
          checkout_types.each do |checkout_type|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
              row.item(:option).value(checkout_type.display_name.localize)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  if t == num_for_last_page - 1
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 2)
                    datas.each do |data|
                      sum += data.value
                    end
                    row.item("value#13").value(sum)
                  end
                end
              end
              if checkout_type == checkout_types.last
                row.item(:library_line).show
                row.item(:condition_line).show
                line_for_items(row) if library.shelves.size < 1
              end
            end
          end
          # accept items each shelves and call_numbers
          library.shelves.each do |shelf|
            report.page.list(:list).add_row do |row|
              row.item(:library).value("(#{shelf.display_name})")
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => nil, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else 
                 num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => nil, :option => 2).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  if t == num_for_last_page - 1
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => nil, :option => 2)
                    datas.each do |data|
                      sum += data.value
                    end
                    row.item("value#13").value(sum)
                  end
                end
              end
              row.item("library_line").show
              row.item("condition_line").show
              line_for_items(row) if shelf == library.shelves.last && call_numbers.nil?
            end
            unless call_numbers.nil?
              call_numbers.each do |num|
                report.page.list(:list).add_row do |row|
                  row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                  row.item(:option).value(num)
                  if start_date != 27
                    13.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 2).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                    end
                  else
                    num_for_last_page.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 2).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                      if t == num_for_last_page - 1
                        sum = 0
                        datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 2)
                        datas.each do |data|
                          sum += data.value
                        end
                        row.item("value#13").value(sum)
                      end
                    end
                  end
                  if num == call_numbers.last
                    row.item("library_line").show
                    row.item("condition_line").show
                    line_for_items(row) if shelf == library.shelves.last
                  end
                end
              end
            end
          end
        end

        # remove items
        report.start_new_page
        report.page.item(:date).value(Time.now)
        report.page.item(:year).value(term[0,4])
        report.page.item(:month).value(term[4,6])        
        report.page.item(:inout_type).value(I18n.t('statistic_report.remove'))
        # header
        if start_date != 27
          13.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
        else
          num_for_last_page.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
          report.page.list(:list).header.item("column#13").value(I18n.t('statistic_report.sum'))
        end
        # remove items all libraries
        data_type = 211
        report.page.list(:list).add_row do |row|
          row.item(:library).value(I18n.t('statistic_report.all_library'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 3).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 3).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
              if t == num_for_last_page - 1
                sum = 0
                datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :option => 3)
                datas.each do |data|
                  sum += data.value
                end
                row.item("value#13").value(sum)
              end
            end
          end
          row.item(:condition_line).show
        end
        # remove items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
              row.item(:option).value(num)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :call_number => num, :option => 3).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :call_number => num, :option => 3).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  if t == num_for_last_page - 1
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :call_number => num, :option => 3)
                    datas.each do |data|
                      sum += data.value
                    end
                    row.item("value#13").value(sum)
                  end
                end
              end
              row.item("condition_line").show if num == call_numbers.last
            end  
          end
        end
        # remove items each checkout_types
        checkout_types.each do |checkout_type|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
            row.item(:option).value(checkout_type.display_name.localize)
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                if t == num_for_last_page - 1
                  sum = 0
                  datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 3)
                  datas.each do |data|
                    sum += data.value
                  end
                  row.item("value#13").value(sum)
                end
              end
            end  
            row.item("condition_line").show if checkout_type == checkout_types.last
            line_for_items(row) if checkout_type == checkout_types.last
          end
        end
        # remove items each libraries
        libraries.each do |library|
          report.page.list(:list).add_row do |row|
            row.item(:library).value(library.display_name)
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 3).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :option => 3).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                if t == num_for_last_page - 1
                  sum = 0
                  datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :option => 3)
                  datas.each do |data|
                    sum += data.value
                  end
                  row.item("value#13").value(sum)
                end
              end
            end
            row.item(:condition_line).show
          end
          # remove items each call_numbers
          unless call_numbers.nil?
            call_numbers.each do |num|
              report.page.list(:list).add_row do |row|
                row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                row.item(:option).value(num)
                if start_date != 27
                  13.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :call_number => num, :option => 3).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                  end
                else
                  num_for_last_page.times do |t|
                    value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :call_number => num, :option => 3).first.value rescue 0
                    row.item("value##{t+1}").value(to_format(value))
                    if t == num_for_last_page - 1
                      sum = 0
                      datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :call_number => num, :option => 3)
                      datas.each do |data|
                        sum += data.value
                      end
                      row.item("value#13").value(sum)
                    end
                  end
                end
                row.item("condition_line").show if num == call_numbers.last
              end
            end
          end
          # remove items each checkout_types
          checkout_types.each do |checkout_type|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
              row.item(:option).value(checkout_type.display_name.localize)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  if t == num_for_last_page - 1
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 3)
                    datas.each do |data|
                      sum += data.value
                    end
                    row.item("value#13").value(sum)
                  end
                end
              end
              if checkout_type == checkout_types.last
                row.item(:library_line).show
                row.item(:condition_line).show
                line_for_items(row) if library.shelves.size < 1
              end
            end
          end
          # remove items each shelves and call_numbers
          library.shelves.each do |shelf|
            report.page.list(:list).add_row do |row|
              row.item(:library).value("(#{shelf.display_name})")
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => nil, :option => 3).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else 
                 num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => nil, :option => 3).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  if t == num_for_last_page - 1
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => nil, :option => 3)
                    datas.each do |data|
                      sum += data.value
                    end
                    row.item("value#13").value(sum)
                  end
                end
              end
              row.item("library_line").show
              row.item("condition_line").show
              line_for_items(row) if shelf == library.shelves.last && call_numbers.nil?
            end
            unless call_numbers.nil?
              call_numbers.each do |num|
                report.page.list(:list).add_row do |row|
                  row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                  row.item(:option).value(num)
                  if start_date != 27
                    13.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 3).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                    end
                  else
                    num_for_last_page.times do |t|
                      value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 3).first.value rescue 0
                      row.item("value##{t+1}").value(to_format(value))
                      if t == num_for_last_page - 1
                        sum = 0
                        datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 3)
                        datas.each do |data|
                          sum += data.value
                        end
                        row.item("value#13").value(sum)
                      end
                    end
                  end
                  if num == call_numbers.last
                    row.item("library_line").show
                    row.item("condition_line").show
                    line_for_items(row) if shelf == library.shelves.last
                  end
                end
              end
            end
          end
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_inout_daily_tsv(term)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_inout_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:library, 'statistic_report.library'],
      [:shelf, 'activerecord.models.shelf'],
      [:condition, 'statistic_report.condition'],
      [:option, 'statistic_report.option'] 
    ]
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    days = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      days.times do |t|
        row << I18n.t('statistic_report.date', :num => t+1)
        columns << ["#{term}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      data_type = 211
      # accept items all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :library
          row << I18n.t('statistic_report.all_library')
        when :shelf
          row << ""
        when :condition
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => 0, :option => 2).first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"    
      # accept items each call_numbers
      unless call_numbers.nil?
        call_numbers.each do |num|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :library
              row << I18n.t('statistic_report.all_library')
            when :shelf
              row << ""
            when :condition
              row << I18n.t('activerecord.attributes.item.call_number')
            when :option
              row << num
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => 0, :call_number => num, :option => 2).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # accept items each checkout_types
      checkout_types.each do |checkout_type|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :library
            row << I18n.t('statistic_report.all_library')
          when :shelf
            row << ""
          when :condition
            row << I18n.t('activerecord.models.checkout_type')
          when :option
            row << checkout_type.display_name.localize
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # accept items each libraries
      libraries.each do |library|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :library
            row << library.display_name.localize
          when :shelf
            row << ""
          when :condition
            row << ""
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => library.id, :option => 2).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # accept items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            row = []
            sum = 0
            columns.each do |column|
              case column[0]
              when :library
                row << library.display_name.localize
              when :shelf
                row << ""
              when :condition
                row << I18n.t('activerecord.attributes.item.call_number')
              when :option
                row << num
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => library.id, :call_number => num, :option => 2).first.value rescue 0
                sum += value
                row << to_format(value)
              end
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
          # accept items each checkout_types
          checkout_types.each do |checkout_type|
            row = []
            sum = 0
            columns.each do |column|
              case column[0]
              when :library
                row << library.display_name.localize
              when :shelf
                row << ""
              when :condition
                row << I18n.t('activerecord.models.checkout_type')
              when :option
                row << checkout_type.display_name.localize
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
                sum += value
                row << to_format(value)
              end
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
          # accept items each shelves and call_numbers
          library.shelves.each do |shelf|
            row = []
            sum = 0
            columns.each do |column|
              case column[0]
              when :library
                row << library.display_name.localize
              when :shelf
                row << shelf.display_name.localize
              when :condition
                row << ""
              when :option
                row << ""
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => nil, :option => 2).first.value rescue 0
                sum += value
                row << to_format(value)
              end
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
            unless call_numbers.nil?
              call_numbers.each do |num|
                row = []
                sum = 0
                columns.each do |column|
                  case column[0]
                  when :library
                    row << library.display_name.localize
                  when :shelf
                    row << shelf.display_name.localize
                  when :condition
                    row << I18n.t('activerecord.attributes.item.call_number')
                  when :option
                    row << num
                  when "sum"
                    row << to_format(sum)
                  else
                    value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 2).first.value rescue 0
                    sum += value
                    row << to_format(value)
                  end
                end
                output.print "\""+row.join("\"\t\"")+"\"\n"
              end
            end
          end
        end
      end
    end
    return tsv_file
  end

  def self.get_inout_monthly_pdf(term)
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    begin 
      report = ThinReports::Report.new :layout => get_layout_path("inout_items_monthly")

      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      report.start_new_page
      report.page.item(:date).value(Time.now)       
      report.page.item(:term).value(term)

      # accept items
      report.page.item(:inout_type).value(I18n.t('statistic_report.accept'))
      
      # accept items all libraries
      data_type = 111
      report.page.list(:list).add_row do |row|
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        12.times do |t|
          if t < 3 # for Japanese fiscal year
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 2).first.value rescue 0
          else
            value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 2).first.value rescue 0
          end
          row.item("value#{t+1}").value(to_format(value))
          sum += value
        end  
        row.item("valueall").value(sum)
        row.item("condition_line").show
      end
      # accept items each call_numbers
      unless call_numbers.nil?
        call_numbers.each do |num|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
            row.item(:option).value(num)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :call_number => num, :option => 2).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :call_number => num, :option => 2).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end  
            row.item("valueall").value(sum)
            row.item("condition_line").show if num == call_numbers.last
          end
        end
      end
      # accept items each checkout_types
      checkout_types.each do |checkout_type|
        report.page.list(:list).add_row do |row|
          row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
          row.item(:option).value(checkout_type.display_name.localize)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end  
          row.item("valueall").value(sum)
          row.item("condition_line").show if checkout_type == checkout_types.last
          line_for_items(row) if checkout_type == checkout_types.last
        end
      end
      # accept items each library
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 2).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 2).first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end  
          row.item("valueall").value(sum)
          row.item("condition_line").show
        end
        # accept items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
              row.item(:option).value(num)
              sum = 0
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :call_number => num, :option => 2).first.value rescue 0
                else
                  value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :call_number => num, :option => 2).first.value rescue 0
                end
                row.item("value#{t+1}").value(to_format(value))
                sum += value
              end
              row.item("valueall").value(sum)
              row.item("condition_line").show if num == call_numbers.last
            end
          end
        end
        # accept items each checkout_types
        checkout_types.each do |checkout_type|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
            row.item(:option).value(checkout_type.display_name.localize)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            if checkout_type == checkout_types.last
              row.item(:library_line).show
              row.item(:condition_line).show
              line_for_items(row) if library.shelves.size < 1
            end
          end
        end
        # accept items each shelves and call_numbers
        library.shelves.each do |shelf|
          report.page.list(:list).add_row do |row|
            row.item(:library).value("(#{shelf.display_name})")
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :option => 2).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :option => 2).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item("library_line").show
            row.item("condition_line").show
            line_for_items(row) if shelf == library.shelves.last && call_numbers.nil?
          end
          unless call_numbers.nil?
            call_numbers.each do |num|
              report.page.list(:list).add_row do |row|
                row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                row.item(:option).value(num)
                sum = 0
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 2).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 2).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  sum += value
                end
                row.item("valueall").value(sum)
                if num == call_numbers.last
                  row.item("library_line").show
                  row.item("condition_line").show
                  line_for_items(row) if shelf == library.shelves.last
                end
              end
            end
          end
        end
      end

      # remove items
      report.start_new_page
      report.page.item(:date).value(Time.now)       
      report.page.item(:term).value(term)
      report.page.item(:inout_type).value(I18n.t('statistic_report.remove'))
      
      # remove items all libraries
      data_type = 111
      report.page.list(:list).add_row do |row|
        row.item(:library).value(I18n.t('statistic_report.all_library'))
        sum = 0
        12.times do |t|
          if t < 3 # for Japanese fiscal year
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 3).first.value rescue 0
          else
            value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 3).first.value rescue 0
          end
          row.item("value#{t+1}").value(to_format(value))
          sum += value
        end  
        row.item("valueall").value(sum)
        row.item("condition_line").show
      end
      # remove items each call_numbers
      unless call_numbers.nil?
        call_numbers.each do |num|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
            row.item(:option).value(num)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :call_number => num, :option => 3).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :call_number => num, :option => 3).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end  
            row.item("valueall").value(sum)
            row.item("condition_line").show if num == call_numbers.last
          end
        end
      end
      # remove items each checkout_types
      checkout_types.each do |checkout_type|
        report.page.list(:list).add_row do |row|
          row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
          row.item(:option).value(checkout_type.display_name.localize)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end  
          row.item("valueall").value(sum)
          row.item("condition_line").show if checkout_type == checkout_types.last
          line_for_items(row) if checkout_type == checkout_types.last
        end
      end
      # remove items each library
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 3).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :option => 3).first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end  
          row.item("valueall").value(sum)
          row.item("condition_line").show
        end
        # remove items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            report.page.list(:list).add_row do |row|
              row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
              row.item(:option).value(num)
              sum = 0
              12.times do |t|
                if t < 3 # for Japanese fiscal year
                  datas = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :call_number => num, :option => 3)
                else
                  datas = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :call_number => num, :option => 3)
                end
                value = 0
                datas.each do |data|
                  value += data.value
                end
                row.item("value#{t+1}").value(to_format(value))
                sum += value
              end
              row.item("valueall").value(sum)
              row.item("condition_line").show if num == call_numbers.last
            end
          end
        end
        # remove items each checkout_types
        checkout_types.each do |checkout_type|
          report.page.list(:list).add_row do |row|
            row.item(:condition).value(I18n.t('activerecord.models.checkout_type')) if checkout_type == checkout_types.first 
            row.item(:option).value(checkout_type.display_name.localize)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            if checkout_type == checkout_types.last
              row.item(:library_line).show
              row.item(:condition_line).show
              line_for_items(row) if library.shelves.size < 1
            end
          end
        end
        # remove items each shelves and call_numbers
        library.shelves.each do |shelf|
          report.page.list(:list).add_row do |row|
            row.item(:library).value("(#{shelf.display_name})")
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                datas = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :option => 3)
              else
                datas = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :option => 3)
              end
              value = 0
              datas.each do |data|
                value += data.value
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item("library_line").show
            row.item("condition_line").show
            line_for_items(row) if shelf == library.shelves.last && call_numbers.nil?
          end
          unless call_numbers.nil?
            call_numbers.each do |num|
              report.page.list(:list).add_row do |row|
                row.item(:condition).value(I18n.t('activerecord.attributes.item.call_number')) if num == call_numbers.first 
                row.item(:option).value(num)
                sum = 0
                12.times do |t|
                  if t < 3 # for Japanese fiscal year
                    value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 3).first.value rescue 0
                  else
                    value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 3).first.value rescue 0
                  end
                  row.item("value#{t+1}").value(to_format(value))
                  sum += value
                end
                row.item("valueall").value(sum)
                if num == call_numbers.last
                  row.item("library_line").show
                  row.item("condition_line").show
                  line_for_items(row) if shelf == library.shelves.last
                end
              end
            end
          end
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_inout_monthly_tsv(term)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_inout_report.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:library, 'statistic_report.library'],
      [:shelf, 'activerecord.models.shelf'],
      [:condition, 'statistic_report.condition'],
      [:option, 'statistic_report.option']
    ]
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      9.times do |t|
        row << I18n.t('statistic_report.month', :num => t+4)
        columns << ["#{term}#{"%02d" % (t + 4)}"]
      end
      3.times do |t|
        row << I18n.t('statistic_report.month', :num => t+1)
        columns << ["#{term.to_i + 1}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # accept items all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.accept')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :shelf
          row << ""
        when :condition
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => 0, :option => 2).first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # accept items each call_numbers
      unless call_numbers.nil?
        call_numbers.each do |num|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.accept')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :shelf
               row << ""
            when :condition
              row << I18n.t('activerecord.attributes.item.call_number')
            when :option
              row << num
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => 0, :call_number => num, :option => 2).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # accept items each checkout_types
      checkout_types.each do |checkout_type|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.accept')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :shelf
            row << ""
          when :condition
            row << I18n.t('activerecord.models.checkout_type')
          when :option
            row << checkout_type.display_name.localize
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # accept items each library
      libraries.each do |library|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.accept')
          when :shelf
            row << ""
          when :library
            row << library.display_name.localize
          when :condition
            row << ""
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :option => 2).first.value rescue 0 
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # accept items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            row = []
            sum = 0
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.accept')
              when :library
                row << library.display_name.localize
              when :shelf
                row << ""
              when :condition
                row << I18n.t('activerecord.attributes.item.call_number')
              when :option
                row << num
              when "sum"
                row << to_format(sum)
              else
                value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :call_number => num, :option => 2).first.value rescue 0
                sum += value
                row << to_format(value)
              end
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
        # accept items each checkout_types
        checkout_types.each do |checkout_type|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.accept')
            when :library
              row << library.display_name.localize
            when :shelf
              row << ""
            when :condition
              row << I18n.t('activerecord.models.checkout_type')
            when :option
              row << checkout_type.display_name.localize
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 2).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # accept items each shelves and call_numbers
        library.shelves.each do |shelf|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.accept')
            when :library
              row << library.display_name.localize
            when :shelf
              row << shelf.display_name.localize
            when :condition
              row << ""
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :shelf_id => shelf.id, :option => 2).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless call_numbers.nil?
            call_numbers.each do |num|
              row = []
              sum = 0
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.accept')
                when :library
                  row << library.display_name.localize
                when :shelf
                  row << shelf.display_name.localize
                when :condition
                  row << I18n.t('activerecord.attributes.item.call_number')
                when :option
                  row << num
                when "sum"
                  row << to_format(sum)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 2).first.value rescue 0
                  sum += value
                  row << to_format(value)
                end
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
        end
      end
      # remove items all libraries
      row = []
      sum = 0
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.remove')
        when :library
          row << I18n.t('statistic_report.all_library')
        when :shelf
          row << ""
        when :condition
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => 0, :option => 3).first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # remove items each call_numbers
      unless call_numbers.nil?
        call_numbers.each do |num|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.remove')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :shelf
              row << ""
            when :condition
              row << I18n.t('activerecord.attributes.item.call_number')
            when :option
              row << num
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => 0, :call_number => num, :option => 3).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # remove items each checkout_types
      checkout_types.each do |checkout_type|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.remove')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :shelf
            row << ""
          when :condition
            row << I18n.t('activerecord.models.checkout_type')
          when :option
            row << checkout_type.display_name.localize
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => 0, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # remove items each library
      libraries.each do |library|
        row = []
        sum = 0
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.remove')
          when :library
            row << library.display_name.localize
          when :shelf
            row << ""
          when :condition
            row << ""
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :option => 3).first.value rescue 0 
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # remove items each call_numbers
        unless call_numbers.nil?
          call_numbers.each do |num|
            row = []
            sum = 0
            columns.each do |column|
              case column[0]
              when :type
                row << I18n.t('statistic_report.remove')
              when :library
                row << library.display_name.localize
              when :shelf
                row << ""
              when :condition
                row << I18n.t('activerecord.attributes.item.call_number')
              when :option
                row << num
              when "sum"
                row << to_format(sum)
              else
                datas = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :call_number => num, :option => 3)
                value = 0
                datas.each do |data|
                  value += data.value
                end
                sum += value
                row << to_format(value)
              end
            end
            output.print "\""+row.join("\"\t\"")+"\"\n"
          end
        end
        # remove items each checkout_types
        checkout_types.each do |checkout_type|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.remove')
            when :library
              row << library.display_name.localize
            when :shelf
              row << ""
            when :condition
              row << I18n.t('activerecord.models.checkout_type')
            when :option
              row << checkout_type.display_name.localize
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :checkout_type_id => checkout_type.id, :option => 3).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # remove items each shelves and call_numbers
        library.shelves.each do |shelf|
          row = []
          sum = 0
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.remove')
            when :library
              row << library.display_name.localize
            when :shelf
              row << shelf.display_name.localize
            when :condition
              row << ""
            when :option
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = 0
              datas = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :shelf_id => shelf.id, :option => 3)
              datas.each do |data|
                value += data.value
              end
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
          unless call_numbers.nil?
            call_numbers.each do |num|
              row = []
              sum = 0
              columns.each do |column|
                case column[0]
                when :type
                  row << I18n.t('statistic_report.remove')
                when :library
                  row << library.display_name.localize
                when :shelf
                  row << shelf.display_name.localize
                when :condition
                  row << I18n.t('activerecord.attributes.item.call_number')
                when :option
                  row << num
                when "sum"
                  row << to_format(sum)
                else
                  value = Statistic.where(:yyyymm => column[0], :data_type => 111, :library_id => library.id, :shelf_id => shelf.id, :call_number => num, :option => 3).first.value rescue 0
                  sum += value
                  row << to_format(value)
                end
              end
              output.print "\""+row.join("\"\t\"")+"\"\n"
            end
          end
        end
      end
    end
    return tsv_file
  end

  def self.get_loans_daily_pdf(term)
    logger.error "create daily inter library loans statistic report: #{term}"
    libraries = Library.all
    begin
      report = ThinReports::Report.new :layout => get_layout_path("loans_daily")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      num_for_last_page = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i - 26
      libraries.each do |library|
        [1,14,27].each do |start_date| # for 3 pages
          report.start_new_page
          report.page.item(:date).value(Time.now)
          report.page.item(:year).value(term[0,4])
          report.page.item(:month).value(term[4,6])        
          report.page.item(:library).value(library.display_name.localize)        
          # header
          if start_date != 27
            13.times do |t|
              report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
            end
          else
            num_for_last_page.times do |t|
              report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
            end
            report.page.list(:list).header.item("column#13").value(I18n.t('statistic_report.sum'))
          end
          # checkout loan
          data_type = 261
          libraries.each do |borrowing_library|
            next if library == borrowing_library
            report.page.list(:list).add_row do |row|
              row.item(:loan_type).value(I18n.t('statistic_report.checkout_loan')) if borrowing_library == libraries.first || (borrowing_library == libraries[1] && library == libraries.first)
              row.item(:borrowing_library).value(borrowing_library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  if t == num_for_last_page - 1
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id)
                    datas.each do |data|
                      sum += data.value
                    end
                    row.item("value#13").value(sum)
                  end
                end
              end
              if borrowing_library == libraries.last || (borrowing_library == libraries[-2] && library == libraries.last)
                line_loan(row)
              end
            end
          end
          # checkin loan
          data_type = 262
          libraries.each do |borrowing_library|
            next if library == borrowing_library
            report.page.list(:list).add_row do |row|
              row.item(:loan_type).value(I18n.t('statistic_report.checkin_loan')) if borrowing_library == libraries.first || (borrowing_library == libraries[1] && library == libraries.first)
              row.item(:borrowing_library).value(borrowing_library.display_name)
              if start_date != 27
                13.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                end
              else
                num_for_last_page.times do |t|
                  value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
                  row.item("value##{t+1}").value(to_format(value))
                  if t == num_for_last_page - 1
                    sum = 0
                    datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id)
                    datas.each do |data|
                      sum += data.value
                    end
                    row.item("value#13").value(sum)
                  end
                end
              end
              row.item(:type_line).show if borrowing_library == libraries.last
            end
          end
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_loans_daily_tsv(term)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_loans_daily.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:library, 'statistic_report.library'],
      [:loan_type, 'statistic_report.loan_type'],
      [:loan_library, 'statistic_report.loan_library'] 
    ]
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    days = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      days.times do |t|
        row << I18n.t('statistic_report.date', :num => t+1)
        columns << ["#{term}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"
      libraries.each do |library|
        # checkout loan
        data_type = 261
        libraries.each do |borrowing_library|
          next if library == borrowing_library
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :library
              row << library.display_name.localize
            when :loan_type
              row << I18n.t('statistic_report.checkout_loan')
            when :loan_library
              row << borrowing_library.display_name.localize
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # checkin loan
        data_type = 262
        libraries.each do |borrowing_library|
          next if library == borrowing_library
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :library
              row << library.display_name.localize
            when :loan_type
              row << I18n.t('statistic_report.checkin_loan')
            when :loan_library
              row << borrowing_library.display_name.localize
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymmdd => column[0], :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
    end
    return tsv_file
  end

  def self.get_loans_monthly_pdf(term)
    logger.error "create monthly inter library loans statistic report: #{term}"
    libraries = Library.all
    begin
      report = ThinReports::Report.new :layout => get_layout_path("loans_monthly")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      libraries.each do |library|
        report.start_new_page
        report.page.item(:date).value(Time.now)
        report.page.item(:term).value(term)
        report.page.item(:library).value(library.display_name.localize)        
        # checkout loan
        data_type = 161
        libraries.each do |borrowing_library|
          next if library == borrowing_library
          report.page.list(:list).add_row do |row|
            row.item(:loan_type).value(I18n.t('statistic_report.checkout_loan')) if borrowing_library == libraries.first || (borrowing_library == libraries[1] && library == libraries.first)
            row.item(:borrowing_library).value(borrowing_library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term.to_i}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            line_loan(row) if borrowing_library == libraries.last || (borrowing_library == libraries[-2] && library == libraries.last)
          end
        end
        # checkin loan
        data_type = 162
        libraries.each do |borrowing_library|
          next if library == borrowing_library
          report.page.list(:list).add_row do |row|
            row.item(:loan_type).value(I18n.t('statistic_report.checkin_loan')) if borrowing_library == libraries.first || (borrowing_library == libraries[1] && library == libraries.first)
            row.item(:borrowing_library).value(borrowing_library.display_name)
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term.to_i}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum += value
            end
            row.item("valueall").value(sum)
            row.item(:type_line).show if borrowing_library == libraries.last || (borrowing_library == libraries[-2] && library == libraries.last)
          end
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_loans_monthly_tsv(term)
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_loans_monthly.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:library, 'statistic_report.library'],
      [:loan_type, 'statistic_report.loan_type'],
      [:loan_library, 'statistic_report.loan_library'] 
    ]
    libraries = Library.all
    checkout_types = CheckoutType.all
    call_numbers = Statistic.call_numbers
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      9.times do |t|
        row << I18n.t('statistic_report.month', :num => t+4)
        columns << ["#{term}#{"%02d" % (t + 4)}"]
      end
      3.times do |t|
        row << I18n.t('statistic_report.month', :num => t+1)
        columns << ["#{term.to_i + 1}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      libraries.each do |library|
        # checkout loan
        data_type = 161
        libraries.each do |borrowing_library|
          next if library == borrowing_library
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :library
              row << library.display_name.localize
            when :loan_type
              row << I18n.t('statistic_report.checkout_loan')
            when :loan_library
              row << borrowing_library.display_name.localize
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        # checkin loan
        data_type = 162
        libraries.each do |borrowing_library|
          next if library == borrowing_library
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :library
              row << library.display_name.localize
            when :loan_type
              row << I18n.t('statistic_report.checkin_loan')
            when :loan_library
              row << borrowing_library.display_name.localize
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => library.id, :borrowing_library_id => borrowing_library.id).first.value rescue 0
              sum += value
              row << to_format(value)
            end
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
    end
    return tsv_file
  end

  def self.get_groups_monthly_pdf(term)
    corporates = User.corporate
    if corporates.blank?
      return false
    end
    dir_base = "#{Rails.root}/private/system"
    begin
      report = ThinReports::Report.new :layout => get_layout_path("groups_monthly")

      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      report.start_new_page
      report.page.item(:date).value(Time.now)       
      report.page.item(:term).value(term)
      # checkout items each corporate users
      corporates.each do |user|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkout_items')) if user == corporates.first
          row.item(:user_name).value(user.agent.full_name)   
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :user_id => user.id).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :user_id => user.id).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          if user == corporates.last
            row.item(:library_line).show 
            line_for_libraries(row)
          end
        end
      end
      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      logger.error $@.join('\n')
      return false
    end	
  end

  def self.get_groups_monthly_tsv(term)
    corporates = User.corporate
    if corporates.blank?
      return false
    end
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_groups_monthly.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:user_name, 'statistic_report.corporate_name']
    ]
    libraries = Library.all
    checkout_types = CheckoutType.all
    user_groups = UserGroup.all
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      9.times do |t|
        row << I18n.t('statistic_report.month', :num => t+4)
        columns << ["#{term}#{"%02d" % (t + 4)}"]
      end
      3.times do |t|
        row << I18n.t('statistic_report.month', :num => t+1)
        columns << ["#{term.to_i + 1}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      corporates = User.corporate
      # checkout items each corporate users
      corporates.each do |user|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :user_name
            row << user.agent.full_name
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 121, :library_id => 0, :user_id => user.id).first.value rescue 0
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
    end
    return tsv_file
  end

  def self.get_groups_daily_pdf(term)
    corporates = User.corporate
    if corporates.blank?
      return false
    end
    begin
      report = ThinReports::Report.new :layout => get_layout_path("groups_daily")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      num_for_last_page = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i - 26
      [1,14,27].each do |start_date| # for 3 pages
        report.start_new_page
        report.page.item(:date).value(Time.now)
        report.page.item(:year).value(term[0,4])
        report.page.item(:month).value(term[4,6])        
        # header
        if start_date != 27
          13.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
        else
          num_for_last_page.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
          report.page.list(:list).header.item("column#13").value(I18n.t('statistic_report.sum'))
        end

        # checkout items each libraries
        corporates.each do |user|
          report.page.list(:list).add_row do |row|
            row.item(:type).value(I18n.t('statistic_report.checkout_items')) if user == corporates.first
            row.item(:user_name).value(user.agent.full_name)   
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 221, :user_id => user.id).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 221, :user_id => user.id).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(:yyyymm => term, :data_type => 221, :user_id => user.id)
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
          end
        end
      end
      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_groups_daily_tsv(term)
    corporates = User.corporate
    if corporates.blank?
      return false
    end
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_groups_daily.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    days = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:user_name, 'statistic_report.corporate_name']
    ]
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      days.times do |t|
        row << I18n.t('statistic_report.date', :num => t+1)
        columns << ["#{term}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      # checkout users each libraries
      corporates.each do |user|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :user_name
            row << user.agent.full_name
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :user_id => user.id).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end  
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
    end
    return tsv_file
  end

  def self.get_departments_monthly_pdf(term)
    libraries = Library.real.all
    departments = Department.all
    manifestation_type_categories = ManifestationType.categories
    user_statuses = UserStatus.all
    if departments.blank?
      return false
    end
    dir_base = "#{Rails.root}/private/system"
    begin
      report = ThinReports::Report.new :layout => get_layout_path("departments_monthly")

      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      report.start_new_page
      report.page.item(:date).value(Time.now)       
      report.page.item(:term).value(term)

      # items all libraries
      data_type = 111
#      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.items'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          row.item(:option).value("#{I18n.t('item.original')}/#{I18n.t('item.copy')}")
          row.item(:option_right).value("#{I18n.t('statistic_report.all')}")
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :manifestation_type_id => 0).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :manifestation_type_id => 0).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
          end  
        end
        # each manifestation type categories
        manifestation_type_categories.each do |c|
          report.page.list(:list).add_row do |row|
            row.item(:option_right).value(I18n.t("manifestation_type.#{c}"))
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(["yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 0", "#{term.to_i + 1}#{"%02d" % (t + 1)}", data_type, 0, ManifestationType.type_ids(c)]).sum(:value) rescue 0
              else
                value = Statistic.where(["yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 0", "#{term}#{"%02d" % (t + 1)}", data_type, 0, ManifestationType.type_ids(c)]).sum(:value) rescue 0
              end 
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            end
            line(row) if manifestation_type_categories.last == c  
          end
        end
#      end
if false
      # spare items
      report.page.list(:list).add_row do |row|
        row.item(:option).value(I18n.t('item.spare'))
        12.times do |t|
          if t < 3 # for Japanese fiscal year
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 4).first.value rescue 0
          else
            value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0, :option => 4).first.value rescue 0
          end
          row.item("value#{t+1}").value(to_format(value))
          row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
        end  
      end
      # each manifestation type categories
      manifestation_type_categories.each do |c|
        report.page.list(:list).add_row do |row|
          row.item(:option_right).value(I18n.t("manifestation_type.#{c}"))
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(["yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 4", "#{term.to_i + 1}#{"%02d" % (t + 1)}", data_type, 0, ManifestationType.type_ids(c)]).first.value rescue 0
            else
              value = Statistic.where(["yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 4", "#{term}#{"%02d" % (t + 1)}", data_type, 0, ManifestationType.type_ids(c)]).first.value rescue 0
            end 
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
          end
          line(row) if manifestation_type_categories.last == c
        end
      end
end

=begin
      # items each library
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.items')) if libraries.size == 1
#          row.item(:library).value(library.display_name)
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
          end  
        end
        # each manifestation type categories
        manifestation_type_categories.each do |c|
          report.page.list(:list).add_row do |row|
            row.item(:option).value(I18n.t("manifestation_type.#{c}"))
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(["yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?)", "#{term.to_i + 1}#{"%02d" % (t + 1)}", data_type, library.id, ManifestationType.type_ids(c)]).first.value rescue 0
              else
                value = Statistic.where(["yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?)", "#{term}#{"%02d" % (t + 1)}", data_type, library.id, ManifestationType.type_ids(c)]).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            end
            if manifestation_type_categories.last == c
              if libraries.last == library
                line(row)
              else
                line_for_libraries(row)
              end
            end
          end
        end
      end
=end
      # open days of each libraries
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.opens')) if libraries.first == library
#          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 113, :library_id => library.id).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 113, :library_id => library.id).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end
      # checkout users all libraries
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkout_users'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => 0, department_id => 0).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => 0, :department_id => 0).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
        end
      end
      # checkout users each libraries
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkout_users')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name.localize)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => library.id).no_condition.first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 122, :library_id => library.id).no_condition.first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
        end
        # checkout users each departments
        departments.each do |department|
          report.page.list(:list).add_row do |row|
            row.item(:department_name).value(department.display_name)   
            sum = 0
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 122, :department_id => department.id).first.value rescue 0
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 122, :department_id => department.id).first.value rescue 0
              end
              row.item("value#{t+1}").value(to_format(value))
              sum = sum + value
            end  
            row.item("valueall").value(sum)
           # row.item(:option_line).show 
            line(row) if libraries.last == library && departments.last == department
          end
        end
      end

      # checkout items all libraries
#      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkout_items'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0).no_condition.first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0).no_condition.first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
        end
#      end
      # checkout items each manifestation type categories
      manifestation_type_categories.each do |c|
        next if c == 'article'
        report.page.list(:list).add_row do |row|
#            row.item(:type).value(I18n.t('statistic_report.checkout_items')) if libraries.size == 1 && department == departments.first
#            row.item(:library).value(library.display_name.localize) if departments.first == department
          row.item(:option).value(I18n.t("manifestation_type.#{c}"))   
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(["yyyymm = ? AND data_type = ? AND manifestation_type_id in (?)", "#{term.to_i + 1}#{"%02d" % (t + 1)}", 121, ManifestationType.type_ids(c)]).sum(:value) rescue 0
            else
              value = Statistic.where(["yyyymm = ? AND data_type = ? AND manifestation_type_id in (?)", "#{term}#{"%02d" % (t + 1)}", 121, ManifestationType.type_ids(c)]).sum(:value) rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value if value
          end  
          row.item("valueall").value(sum)
          row.item(:option_line).show 
          line(row) if manifestation_type_categories.last == c
        end
      end
      # reminder checkout items
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.remind_checkouts'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :option => 5).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => 0, :option => 5).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          row.item(:library_line).show
        end
      end	
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.remind_checkouts')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name.localize)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library.id, :option => 5).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 121, :library_id => library_id, :option => 5).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end
     
      # checkin items
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkin_items'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0).no_condition.first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0).no_condition.first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          row.item(:library_line).show
        end
      end
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkin_items')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library.id).no_condition.first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library.id).no_condition.first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
	  row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end
=begin
      # checkin items remindered
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkin_remindered'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0, :option => 5).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => 0, :option => 5).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          row.item(:library_line).show
        end
      end 
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkin_remindered')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name.localize)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library.id, :option => 5).first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 151, :library_id => library_id, :option => 5).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end
=end
      # reserves all libraries
#      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.reserves'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => 0).no_condition.first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => 0).no_condition.first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          line(row)
        end
#      end
=begin
      # reserves each library
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.reserves')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => library.id).no_condition.first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 133, :library_id => library.id).no_condition.first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          line(row) if libraries.last == library
        end
      end
=end
      # all users all libraries
      data_type = 112
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.users'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          row.item(:option).value(I18n.t('statistic_report.all_users'))
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
          end  
        end
      end
      # all users each libraries
      libraries.each do |library|
        # all users
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.users')) if libraries.size == 1
#          row.item(:library).value(library.display_name)
          row.item(:option).value(I18n.t('statistic_report.all_users'))
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id).no_condition.first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
          end  
        end
        # each departments
        departments.each do |department|
          report.page.list(:list).add_row do |row|
            row.item(:option).value(department.display_name)
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :department_id => department.id).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :department_id => department.id).first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            end
            row.item(:option_line).style(:border_color, '#000000') if departments.last == department 
            row.item(:values_line).style(:border_color, '#000000') if departments.last == department 
          end
        end
        # each user_statuses
        user_statuses.each do |user_status|
          report.page.list(:list).add_row do |row|
            row.item(:option).value(user_status.display_name)
            12.times do |t|
              if t < 3 # for Japanese fiscal year
                value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :user_status_id => user_status.id).first.value rescue 0 
              else
                value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => data_type, :library_id => library.id, :user_status_id => user_status.id).first.value rescue 0 
              end
              row.item("value#{t+1}").value(to_format(value))
              row.item("valueall").value(to_format(value)) if t == 2 # March(end of fiscal year)
            end
            line(row) if libraries.last == library && user_statuses.last == user_status  
          end
        end
      end
      # questions all libraries
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.questions'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 143, :library_id => 0).no_condition.first.value rescue 0
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 143, :library_id => 0).no_condition.first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          row.item(:library_line).show
        end
      end
      # questions each library
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.questions')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 143, :library_id => library.id).no_condition.first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 143, :library_id => library.id).no_condition.first.value rescue 0 
            end
            row.item("value#{t+1}").value(to_format(value))
            sum = sum + value
          end  
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end
      # visiters all libraries
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.visiters'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 116, :library_id => 0).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 116, :library_id => 0).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end
          row.item("valueall").value(sum)
          row.item(:library_line).show
        end
      end
      # visiters of each libraries
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.visiters')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 116, :library_id => library.id).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 116, :library_id => library.id).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end
      # consultations all libraries
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.consultations'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 114, :library_id => 0).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 114, :library_id => 0).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end
          row.item("valueall").value(sum)
          row.item(:library_line).show
        end
      end
      # consultations of each libraries
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.consultations')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 114, :library_id => library.id).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 114, :library_id => library.id).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end
      # copies all libraries
      if libraries.size > 1
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.copies'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 115, :library_id => 0).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 115, :library_id => 0).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end
          row.item("valueall").value(sum)
          row.item(:library_line).show
        end
      end
      # copies of each libraries
      libraries.each do |library|
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.copies')) if libraries.size == 1 && libraries.first == library
#          row.item(:library).value(library.display_name)
          sum = 0
          12.times do |t|
            if t < 3 # for Japanese fiscal year
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}#{"%02d" % (t + 1)}", :data_type => 115, :library_id => library.id).first.value rescue 0 
            else
              value = Statistic.where(:yyyymm => "#{term}#{"%02d" % (t + 1)}", :data_type => 115, :library_id => library.id).first.value rescue 0
            end
            row.item("value#{t+1}").value(to_format(value))
            sum += value
          end
          row.item("valueall").value(sum)
          row.item(:library_line).show
          line(row) if library == libraries.last
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      logger.error $@.join('\n')
      return false
    end	
  end

  def self.get_departments_monthly_tsv(term)
    departments = Department.all
    if departments.blank?
      return false
    end
    manifestation_type_categories = ManifestationType.categories
    user_statuses = UserStatus.all
    libraries = Library.real.all
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_departments_monthly.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:manifestation_type, 'statistic_report.manifestation_type'],      
      [:department_name, 'statistic_report.department_name'],
      [:option, 'statistic_report.option']
    ]
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      9.times do |t|
        row << I18n.t('statistic_report.month', :num => t+4)
        columns << ["#{term}#{"%02d" % (t + 4)}"]
      end
      3.times do |t|
        row << I18n.t('statistic_report.month', :num => t+1)
        columns << ["#{term.to_i + 1}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      corporates = User.corporate

      # items
      data_type = 111
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.items')
        when :option
          row << ""
        when :manifestation_type
          row << I18n.t('statistic_report.all')
        when :department_name 
          row << ""
        when "sum"
          value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
          row << to_format(value)
        else
          value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      manifestation_type_categories.each do |c|
        row = []  
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.items')
          when :option
            row << ""
          when :manifestation_type
            row << I18n.t("manifestation_type.#{c}")
          when :department_name 
            row << ""
          when "sum"
            value = Statistic.where("yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?)", "#{term.to_i + 1}03", data_type, 0, ManifestationType.type_ids(c)).sum(:value) rescue 0
            row << to_format(value)
          else
            value = Statistic.where("yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?)", column[0], data_type, 0, ManifestationType.type_ids(c)).sum(:value) rescue 0
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # spare items
if false
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.items')
        when :option
          row << I18n.t('item.spare')
        when :manifestation_type
          row << ""
        when :department_name 
          row << ""
        when "sum"
          value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => data_type, :library_id => 0, :option => 4).first.value rescue 0
          row << to_format(value)
        else
          value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :option => 4).first.value rescue 0
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"
      manifestation_type_categories.each do |c|
        row = []  
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.items')
          when :option
            row << I18n.t('item.spare')
          when :manifestation_type
            row << I18n.t("manifestation_type.#{c}")
          when :department_name 
            row << ""
          when "sum"
            value = Statistic.where("yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 4", "#{term.to_i + 1}03", data_type, 0, ManifestationType.type_ids(c)).first.value rescue 0
            row << to_format(value)
          else
            value = Statistic.where("yyyymm = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 4", column[0], data_type, 0, ManifestationType.type_ids(c)).first.value rescue 0
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
end

      # open days of each libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.opens')
          when :library
            row << library.display_name
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name 
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 113, :library_id => library.id).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout users all libraries
      data_type = 122
#     if libraries.size > 1
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name 
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # checkout users each departments
        departments.each do |department|
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_users')
            when :option
              row << ""
            when :manifestation_type
              row << ""
            when :department_name
              row << department.display_name
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 122, :department_id => department.id).first.value rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
#     end
      # checkout items all libraries
      data_type = 121
#     if libraries.size > 1
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # checkout items each manifestation type categories
        manifestation_type_categories.each do |c|
          next if c == 'article'
          sum = 0
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.checkout_items')
            when :option
              row << ""
            when :manifestation_type
              row << I18n.t("manifestation_type.#{c}")
            when :department_name
              row << ""
            when "sum"
              row << to_format(sum)
            else
              value = Statistic.where(["yyyymm = ? AND data_type = ? AND manifestation_type_id in (?)", column[0], 121, ManifestationType.type_ids(c)]).sum(:value) rescue 0
              sum += value
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
#     end
      # remind checkout items
#      if libraries.size > 1
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.remind_checkouts')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
#      end
      # checkin items
#      if libraries.size > 1
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkin_items')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 151, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
#      end
=begin
      # checkin items remindered
#      if libraries.size > 1
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkin_remindered')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 151, :library_id => 0, :option => 5).first.value rescue 0
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
#      end
=end
      # reserves all libraries
#      if libraries.size > 1
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.reserves')  
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when :department_name
            row << ""
          when :manifestation_type
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 133, :library_id => 0).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
#      end
      # all users
      libraries.each do |library|
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.users')
          when :library
            row << I18n.t('statistic_report.all_library')
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            value = Statistic.where(:yyyymm => "#{term.to_i + 1}03", :data_type => 112, :library_id => library.id).no_condition.first.value rescue 0
            row << to_format(value)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => library.id).no_condition.first.value rescue 0
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
        # users each departments
        departments.each do |department|
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.users')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << ""
            when :manifestation_type
              row << ""
            when :department_name
              row << department.display_name
            when "sum"
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => library.id, :department_id => department.id).first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => library.id, :department_id => department.id).first.value rescue 0
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
        user_statuses.each do |user_status|
          row = []
          columns.each do |column|
            case column[0]
            when :type
              row << I18n.t('statistic_report.users')
            when :library
              row << I18n.t('statistic_report.all_library')
            when :option
              row << user_status.display_name
            when :manifestation_type
              row << ""
            when :department_name
              row << ""
            when "sum"
              value = Statistic.where(:yyyymm => "#{term.to_i + 1}03}", :data_type => 112, :library_id => library.id, :user_status_id => user_status.id).first.value rescue 0
              row << to_format(value)
            else
              value = Statistic.where(:yyyymm => column[0], :data_type => 112, :library_id => library.id, :user_status_id => user_status.id).first.value rescue 0
              row << to_format(value)
            end  
          end
          output.print "\""+row.join("\"\t\"")+"\"\n"
        end
      end
      # questions each library
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.questions')
          when :library
            row << library.display_name
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 143, :library_id => library.id).no_condition.first.value rescue 0 
            sum += value
            row << to_format(value)
          end
        end  
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # visiters
#      if libraries.size > 1
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.visiters')
          when :library
            row << "" # library.display_name.localize
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 116, :library_id => 0).first.value rescue 0 
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
#      end
      # consultations
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.consultations')
          when :library
            row << library.display_name
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 114, :library_id => library.id).first.value rescue 0 
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # copies all libraries
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.copies')
          when :library
            row << library.display_name
          when :option
            row << ""
          when :manifestation_type
            row << ""
          when :department_name
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymm => column[0], :data_type => 115, :library_id => library.id).first.value rescue 0 
            sum += value
            row << to_format(value)
          end  
        end
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end

    end
    return tsv_file
  end

  def self.get_departments_daily_pdf(term)
    departments = Department.all
    if departments.blank?
      return false
    end
    manifestation_type_categories = ManifestationType.categories
    begin
      report = ThinReports::Report.new :layout => get_layout_path("department_daily")
      report.events.on :page_create do |e|
        e.page.item(:page).value(e.page.no)
      end
      report.events.on :generate do |e|
        e.pages.each do |page|
          page.item(:total).value(e.report.page_count)
        end
      end

      num_for_last_page = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i - 26
      [1,14,27].each do |start_date| # for 3 pages
        report.start_new_page
        report.page.item(:date).value(Time.now)
        report.page.item(:year).value(term[0,4])
        report.page.item(:month).value(term[4,6])        
        # header
        if start_date != 27
          13.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
        else
          num_for_last_page.times do |t|
            report.page.list(:list).header.item("column##{t+1}").value(I18n.t('statistic_report.date', :num => t+start_date))
          end
          report.page.list(:list).header.item("column#13").value(I18n.t('statistic_report.sum'))
        end
        # items all libraries
        data_type = 211
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.items'))
          row.item(:option).value("#{I18n.t('item.original')}/#{I18n.t('item.copy')}")
          row.item(:option_right).value("#{I18n.t('statistic_report.all')}")
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
              row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
            end
          end
        end
        # each manifestation type categories
        manifestation_type_categories.each do |c|
          report.page.list(:list).add_row do |row|
            row.item(:option_right).value(I18n.t("manifestation_type.#{c}"))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 0", "#{term.to_i}#{"%02d" % (t + start_date)}", data_type, 0, ManifestationType.type_ids(c)]).sum(:value) rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 0", "#{term.to_i}#{"%02d" % (t + start_date)}", data_type, 0, ManifestationType.type_ids(c)]).sum(:value) rescue 0
                row.item("value##{t+1}").value(to_format(value))
                row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
              end
            end
            line(row) if manifestation_type_categories.last == c  
          end
        end
if false
        # spare items 
        report.page.list(:list).add_row do |row|
          row.item(:option).value(I18n.t('item.spare'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 211, :library_id => 0, :option => 4).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 211, :library_id => 0, :option => 4).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
              row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
            end
          end
        end
        manifestation_type_categories.each do |c|
          report.page.list(:list).add_row do |row|
            row.item(:option_right).value(I18n.t("manifestation_type.#{c}"))
            if start_date != 27
              13.times do |t|
                value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 4", "#{term.to_i}#{"%02d" % (t + start_date)}", data_type, 0, ManifestationType.type_ids(c)]).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?) AND option = 4", "#{term.to_i}#{"%02d" % (t + start_date)}", data_type, 0, ManifestationType.type_ids(c)]).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
                row.item("value#13").value(to_format(value)) if t == num_for_last_page - 1
              end
            end
            line(row) if manifestation_type_categories.last == c  
          end
        end
end
        # checkout users all libraries
        data_type = 222
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkout_users'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
            sum = 0
            datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0).no_condition
            datas.each do |data|
              sum = sum + data.value
            end
            row.item("value#13").value(sum)
          end
        end
        # checkout users each departments
        departments.each do |department|
          report.page.list(:list).add_row do |row|
            row.item(:department_name).value(department.display_name)   
            if start_date != 27
              13.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 222, :department_id => department.id).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 222, :department_id => department.id).first.value rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(:yyyymm => term, :data_type => 222, :department_id => department.id)
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
            line(row) if department == departments.last
          end
        end

        # checkout items all libraries
        data_type = 221
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkout_items'))
#          row.item(:library).value(I18n.t('statistic_report.all_library'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_conditionfirst.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
            sum = 0
            datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0).no_condition
            datas.each do |data|
              sum = sum + data.value
            end
            row.item("value#13").value(sum)
          end
        end

        # checkout items each manfiestation type categories
        manifestation_type_categories.each do |c|
          next if c == 'article'
          report.page.list(:list).add_row do |row|
            row.item(:option).value(I18n.t("manifestation_type.#{c}"))   
            if start_date != 27
              13.times do |t|
                value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?)", "#{term.to_i}#{"%02d" % (t + start_date)}", 221, 0, ManifestationType.type_ids(c)]).sum(:value) rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
            else
              num_for_last_page.times do |t|
                value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND library_id = ? AND manifestation_type_id in (?)", "#{term.to_i}#{"%02d" % (t + start_date)}", 221, 0, ManifestationType.type_ids(c)]).sum(:value) rescue 0
                row.item("value##{t+1}").value(to_format(value))
              end
              sum = 0
              datas = Statistic.where(["yyyymm = ? AND data_type = ? AND manifestation_type_id in (?)", term, 221, ManifestationType.type_ids(c)])
              datas.each do |data|
                sum = sum + data.value
              end
              row.item("value#13").value(sum)
            end
            row.item(:library_line).show
            line(row) if manifestation_type_categories.last == c
          end
        end
     
        # remind checkout items
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.remind_checkouts'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
            sum = 0
            datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0, :option => 5)
            datas.each do |data|
              sum = sum + data.value
            end
            row.item("value#13").value(sum)
          end
          line(row)
        end
        # checkin items
        data_type = 251
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkin_items'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
            sum = 0
            datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0).no_condition
            datas.each do |data|
              sum = sum + data.value
            end
            row.item("value#13").value(sum)
          end
          line(row)
        end
=begin
        # checkin items remindered
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.checkin_remindered'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => data_type, :library_id => 0, :option => 5).first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
            sum = 0
            datas = Statistic.where(:yyyymm => term, :data_type => data_type, :library_id => 0).no_condition
            datas.each do |data|
              sum = sum + data.value
            end
            row.item("value#13").value(sum)
          end
          line(row)
        end
=end
        # reserves all libraries
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.reserves'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
          else  
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 233, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
            sum = 0
            datas = Statistic.where(:yyyymm => term, :data_type => 233, :library_id => 0).no_condition
            datas.each do |data|
              sum = sum + data.value
            end
            row.item("value#13").value(sum)
         end
            line(row)
        end  
        # questions all libraries
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.questions'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 243, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end  
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 243, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
            sum = 0              
            datas = Statistic.where(:yyyymm => term, :data_type => 243, :library_id => 0).no_condition
            datas.each do |data|
              sum = sum + data.value 
            end
            row.item("value#13").value(sum)
          end
          line(row)
        end
        # consultations each library
        report.page.list(:list).add_row do |row|
          row.item(:type).value(I18n.t('statistic_report.consultations'))
          if start_date != 27
            13.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 214, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end  
          else
            num_for_last_page.times do |t|
              value = Statistic.where(:yyyymmdd => "#{term.to_i}#{"%02d" % (t + start_date)}", :data_type => 214, :library_id => 0).no_condition.first.value rescue 0
              row.item("value##{t+1}").value(to_format(value))
            end
            sum = 0
            datas = Statistic.where(:yyyymm => term, :data_type => 214, :library_id => 0).no_condition
            datas.each do |data|
              sum = sum + data.value
            end
            row.item("value#13").value(sum)
          end
          line(row)
        end
      end

      return report.generate
    rescue Exception => e
      logger.error "failed #{e}"
      return false
    end
  end

  def self.get_departments_daily_tsv(term)
    departments = Department.all
    if departments.blank?
      return false
    end
    manifestation_type_categories = ManifestationType.categories 
    libraries = Library.real.all
    dir_base = "#{Rails.root}/private/system"
    out_dir = "#{dir_base}/statistic_report/"
    tsv_file = out_dir + "#{term}_departments_daily.tsv"
    FileUtils.mkdir_p(out_dir) unless FileTest.exist?(out_dir)
    days = Time.zone.parse("#{term}01").end_of_month.strftime("%d").to_i
    # header
    columns = [
      [:type,'statistic_report.type'],
      [:manifestation_type, 'statistic_report.manifestation_type'],      
      [:department_name, 'statistic_report.department_name'],
      [:option, 'statistic_report.option']
    ]
    File.open(tsv_file, "w") do |output|
      # add UTF-8 BOM for excel
      output.print "\xEF\xBB\xBF".force_encoding("UTF-8")

      # タイトル行
      row = []
      columns.each do |column|
        row << I18n.t(column[1])
      end
      days.times do |t|
        row << I18n.t('statistic_report.date', :num => t+1)
        columns << ["#{term}#{"%02d" % (t + 1)}"]
      end
      row << I18n.t('statistic_report.sum')
      columns << ["sum"]
      output.print "\""+row.join("\"\t\"")+"\"\n"

      # items all libraries
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.items')
        when :manifestation_type
          row << I18n.t('statistic_report.all')
        when :department_name
          row << ""
        when :option
          row << ""
        when "sum"
          value = Statistic.where(:yyyymmdd => "#{term}#{days}", :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
          row << to_format(value)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => 0).no_condition.first.value rescue 0
          row << to_format(value)
        end
      end  
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # items each manifestation type categories
      manifestation_type_categories.each do |c|
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.items')
          when :manifestation_type
            row << I18n.t("manifestation_type.#{c}")
          when :department_name
            row << ""
          when :option
            row << ""
          when "sum"
            value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND option = 0 AND manifestation_type_id in (?)", "#{term}#{days}", 211, ManifestationType.type_ids(c)]).sum(:value) rescue 0
            row << to_format(value)
          else
            value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND option = 0 AND manifestation_type_id in (?)", column[0], 211, ManifestationType.type_ids(c)]).sum(:value) rescue 0
            row << to_format(value)
          end
        end  
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # spare
if false
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.items')
        when :manifestation_type
          row << ""
        when :department_name
          row << ""
        when :option
          row << I18n.t('item.spare')
        when "sum"
          value = Statistic.where(:yyyymmdd => "#{term}#{days}", :data_type => 211, :library_id => 0, :option=> 4).first.value rescue 0
          row << to_format(value)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 211, :library_id => 0, :option => 4).first.value rescue 0
          row << to_format(value)
        end
      end  
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # items each manifestation type categories
      manifestation_type_categories.each do |c|
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.items')
          when :manifestation_type
            row << I18n.t("manifestation_type.#{c}")
          when :department_name
            row << ""
          when :option
            row << I18n.t('item.spare')
          when "sum"
            value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND option = 4 AND manifestation_type_id in (?)", "#{term}#{days}", 211, ManifestationType.type_ids(c)]).first.value rescue 0
            row << to_format(value)
          else
            value = Statistic.where(["yyyymmdd = ? AND data_type = ? AND option = 4 AND manifestation_type_id in (?)", column[0], 211, ManifestationType.type_ids(c)]).first.value rescue 0
            row << to_format(value)
          end
        end  
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
end
      # checkout users all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkout_users')
        when :manifestation_type
          row << ""
        when :department_name
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 222, :library_id => 0).no_condition.first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end  
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # checkout users each departments
      departments.each do |department|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_users')
          when :department_name
            row << department.display_name
          when :manifestation_type
            row << ""
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 222, :department_id => department.id).first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end  
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout items all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkout_items')
        when :manifestation_type
          row << ""
        when :department_name
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :library_id => 0).no_condition.first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"  
      # checkout items each manifestation_type_categories
      manifestation_type_categories.each do |c|
        next if c == 'article'
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.checkout_items')
          when :department_name
            row << ""
          when :manifestation_type
            row << I18n.t("manifestation_type.#{c}")
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(["yyyymmdd = ? AND library_id = 0 AND data_type = ? AND manifestation_type_id in (?)", column[0], 221, ManifestationType.type_ids(c)]).sum(:value) rescue 0
            sum += value
            row << to_format(value)
          end
        end  
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
      # checkout items reminded
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.remind_checkouts')
        when :department_name
          row << ""
        when :manifestation_type
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 221, :library_id => 0, :option => 5).first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"  
      # checkin items
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkin_items')
        when :department_name
          row << ""
        when :manifestation_type
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 251, :library_id => 0).no_condition.first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end 
      output.print "\""+row.join("\"\t\"")+"\"\n"
=begin
      # checkin items reminded
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.checkin_remindered')
        when :department_name
          row << ""
        when :manifestation_type
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 251, :library_id => 0, :option => 5).first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end
      output.print "\""+row.join("\"\t\"")+"\"\n"  
=end
      # reserves all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.reserves')
        when :department_name
          row << ""
        when :manifestation_type
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 233, :library_id => 0).no_condition.first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end 
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # questions all libraries
      sum = 0
      row = []
      columns.each do |column|
        case column[0]
        when :type
          row << I18n.t('statistic_report.questions')
        when :department_name
          row << ""
        when :manifestation_type
          row << ""
        when :option
          row << ""
        when "sum"
          row << to_format(sum)
        else
          value = Statistic.where(:yyyymmdd => column[0], :data_type => 243, :library_id => 0).no_condition.first.value rescue 0
          sum += value
          row << to_format(value)
        end
      end 
      output.print "\""+row.join("\"\t\"")+"\"\n"
      # consultations each library
      libraries.each do |library|
        sum = 0
        row = []
        columns.each do |column|
          case column[0]
          when :type
            row << I18n.t('statistic_report.consultations')
#          when :library
#            row << library.display_name
          when :department_name
            row << ""
          when :manifestation_type
            row << ""
          when :option
            row << ""
          when "sum"
            row << to_format(sum)
          else
            value = Statistic.where(:yyyymmdd => column[0], :data_type => 214, :library_id => library.id).no_condition.first.value rescue 0
            sum += value
            row << to_format(value)
          end
        end 
        output.print "\""+row.join("\"\t\"")+"\"\n"
      end
#TODO
    end
    return tsv_file
  end

private
  def self.line(row)
    row.item(:type_line).show
    row.item(:library_line).show
    row.item(:library_line).style(:border_color, '#000000')
    row.item(:library_line).style(:border_width, 1)
    row.item(:option_line).style(:border_color, '#000000')
    row.item(:option_line).style(:border_width, 1)
    row.item(:values_line).style(:border_color, '#000000')
    row.item(:values_line).style(:border_width, 1)
  end

  def self.line_loan(row)
    row.item(:type_line).show
    row.item(:type_line).style(:border_color, '#000000')
    row.item(:type_line).style(:border_width, 1)
    row.item(:library_line).show
    row.item(:library_line).style(:border_color, '#000000')
    row.item(:library_line).style(:border_width, 1)
    row.item(:values_line).style(:border_color, '#000000')
    row.item(:values_line).style(:border_width, 1)
  end

  def self.line_for_libraries(row)
    row.item(:library_line).show
    row.item(:library_line).style(:border_color, '#000000')
    row.item(:library_line).style(:border_width, 1)
    row.item(:option_line).style(:border_color, '#000000')
    row.item(:option_line).style(:border_width, 1)
    row.item(:values_line).style(:border_color, '#000000')
    row.item(:values_line).style(:border_width, 1)
  end

  def self.line_for_items(row)
    row.item(:library_line).show
    row.item(:library_line).style(:border_color, '#000000')
    row.item(:library_line).style(:border_width, 1)
    row.item(:condition_line).show
    row.item(:condition_line).style(:border_color, '#000000')
    row.item(:condition_line).style(:border_width, 1)
    row.item(:option_line).style(:border_color, '#000000')
    row.item(:option_line).style(:border_width, 1)
    row.item(:values_line).style(:border_color, '#000000')
    row.item(:values_line).style(:border_width, 1)
  end

  def self.to_format(num = 0)
    num.to_s.gsub(/(\d)(?=(\d{3})+(?!\d))/, '\1,') 
  end

  def self.get_layout_path(filename)
    spec = Gem::Specification.find_by_name("enju_trunk_statistics")
    gem_root = spec.gem_dir
    path = gem_root + "/app/layouts/#{filename}"
    return path
  end

  class GenerateStatisticReportJob
    include Rails.application.routes.url_helpers
    include BackgroundJobUtils

    def initialize(name, target, type, user, options)
      @name     = name
      @target   = target
      @type     = type
      @user     = user
      @options  = options
    end
    attr_accessor :name, :target, :type, :user, :options

    def perform
      user_file = UserFile.new(user)
      StatisticReport.generate_report_internal(target, type, options) do |output|
        io, info = user_file.create(:statisticreport, output.filename)
        if output.result_type == :path
          open(output.path) { |io2| FileUtils.copy_stream(io2, io) }
        else
          io.print output.data
        end
        io.close
        url = my_account_url(:filename => info[:filename], :category => info[:category], :random => info[:random])
        message(
          user,
          I18n.t('statistic_report.output_job_success_subject', :job_name => name),
          I18n.t('statistic_report.output_job_success_body', :job_name => name, :url => url)
        )
      end
    rescue => exception
      message(
        user,
        I18n.t('statistic_report.output_job_error_subject', 
          :job_name => name),
        I18n.t('statistic_report.output_job_error_body', 
          :job_name => name, :message => exception.message+exception.backtrace) 
      )
    end
  end
end
