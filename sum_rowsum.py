class RowSumResult(object):
  """行求和"""
  def __init__(self):
    self.index = -1
    '''求和结果-列下标'''
    self.name = ""
    '''求和结果-列名称'''

    self.item_index_list = []
    '''用于求和-列下标'''
    self.item_name_list = []
    '''用于求和-列名称'''

    self.operator_row_index = -1
    '''用于求和-行下标'''

  def set_index(self, i):
    """求和结果-列下标"""
    self.index = i

  def set_name(self, name):
    """求和结果-列名称"""
    self.name = name

  def set_item_list(self, value: list):
    self.item_index_list = value

  def add_list_item(self, item):
    """求和元素-列下标"""
    if self.item_index_list is None:
      self.item_index_list = []
    self.item_index_list.append(item)

  def add_list_item_name(self, item):
    """求和元素-列名字"""
    if self.item_name_list is None:
      self.item_name_list = []
    self.item_name_list.append(item)

  def get_sumstr(self, helper, row):
    self.operator_row_index = row
    sum_itemlist = []
    for col in self.item_index_list:
      sumstr = helper.get_rowcol_2_str(self.operator_row_index, col)
      sum_itemlist.append(sumstr)
    sumstr = ",".join(sum_itemlist)
    sumstr = f"=SUM({sumstr})"
    return sumstr
