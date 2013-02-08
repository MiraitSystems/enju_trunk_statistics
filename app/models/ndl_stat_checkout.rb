class NdlStatCheckout < ActiveRecord::Base
  belongs_to :ndl_statistic
  attr_accessible :item, :item_type, :user

  item_type_list = ['book', 'magazine', 'other']

  validates_presence_of :item_type, :user, :item
  validates_inclusion_of :item_type, :in => item_type_list
  validates_numericality_of :item, :user
end
