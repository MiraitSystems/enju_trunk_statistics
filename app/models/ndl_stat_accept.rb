class NdlStatAccept < ActiveRecord::Base
  belongs_to :ndl_statistic
  attr_accessible :donation, :production, :purchase, :region, :item_type
  
  item_type_list = ['book', 'magazine', 'other_micro', 'other_av', 'other_file']
  region_list = ['domestic', 'foreign']
  
  validates_presence_of :item_type
  validates_inclusion_of :item_type, :in => item_type_list
  validates_inclusion_of :region, :in => region_list
end
