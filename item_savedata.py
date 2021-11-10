class SaveDataItem():
  """数据持久化单元"""
  def __init__(self, dic) -> None:
    self.set_data(dic)
    self.columnindex = None

  def set_data(self, dic):
    if isinstance(dic, dict):
      for key, value in dic.items():
        self.__dict__[key] = value
    return self

  @staticmethod
  def get_countcell_num(dic):
    '''合计数值'''
    if isinstance(dic, dict):
      return 'countcell' in dic and dic['countcell']
    return None

  @staticmethod
  def get_des_contain_countcell(dic):
    '''获取带合计列的列内容'''
    if isinstance(dic, dict):
      return 'des' in dic and dic['des']
    return None