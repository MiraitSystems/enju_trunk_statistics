class NdlStatJmaPublication < ActiveRecord::Base
  belongs_to :ndl_statistic
  attr_accessible :original_title, :number_string
  
  validates_presence_of :original_title, :number_string
end
