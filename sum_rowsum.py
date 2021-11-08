class RowSumResult(object):
  """行求和"""
  def __init__(self):
    # 求和结果-列下标
    self.index = -1
    # 用于求和-列下标
    self.item_list = []

  def set_index(self, i):
    """求和结果-列下标"""
    self.index = i

  def set_list(self, value: list):
    self.item_list = value

  def add_list_item(self, item):
    """求和元素-列下标"""
    if self.item_list is None:
      self.item_list = []
    self.item_list.append(item)