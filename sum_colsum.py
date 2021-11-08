class ColSumResult(object):
  """
    列求和
    """
  def __init__(self):
    # 求和结果-列下标
    self.result_index = -1
    # 求和对象-列下标
    self.operator_index = -1
    # 求和对象-行开始
    self.rowstart_index = -1
    # 求和对象-行结束
    self.rowend_index = -1

  def set_resultindex(self, i):
    self.result_index = i

  def set_operatorindex(self, i):
    self.operator_index = i

  def set_rowstartindex(self, i):
    self.rowstart_index = i

  def set_rowendindex(self, i):
    self.rowend_index = i

  def get_sumstr(self, helper, startindex, endindex):
    self.set_rowstartindex(startindex)
    self.set_rowendindex(endindex)
    if self.operator_index != -1:
      beginstr = helper.get_rowcol_2_str(self.rowstart_index,
                                         self.operator_index)
      endstr = helper.get_rowcol_2_str(self.rowend_index, self.operator_index)
      result = f"=SUM({beginstr}:{endstr})"
      return result
    return ""