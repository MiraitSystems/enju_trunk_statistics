# -*- encoding: utf-8 -*-
class NdlStatisticsController < ApplicationController
  #load_and_authorize_resource
  before_filter :check_role

  def index
    current_term = Term.current_term.display_name.gsub(/平成(\d+)年度/,'\1').to_i
    @term = current_term unless params[:term]
    respond_to do |format|
      format.html # index.html.erb
    end
  end

  # check role
  def check_role
    unless current_user.try(:has_role?, 'Librarian')
      access_denied; return
    end
  end
  
  # get_ndl_report
  def get_ndl_report
    term = params[:term].strip
    term_name = "平成#{term}年度"

    unless Term.where(:display_name => term_name).exists?
      flash[:message] = t('ndl_report.invalid_term')
      @term = term
      render :index
      return false
    else
      term_id = Term.where(:display_name => term_name).first.id
    end
    unless NdlStatistic.where(:term_id => term_id).exists?
      flash[:message] = t('ndl_report.term_not_found')
      @term = term
      render :index
      return false
    else
      ndl_statistic = NdlStatistic.where(:term_id => term_id).first
      file = NdlStatistic.get_ndl_report_excelx(ndl_statistic)
      send_file file, :filename => "#{term}_#{t('ndl_report.filename_excelx')}",
                      :type => 'application/x-msexcel', :disposition => 'attachment'
    end
  end

end
