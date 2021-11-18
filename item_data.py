class DataItem():
  def __init__(self, date="") -> None:
    self.date = date

  @property
  def date(self):
    return self.__date

  @date.setter
  def date(self, value):
    self.__date = value

  def set_data(self, name, value=""):
    self.__dict__[name] = value

  def set_des(self, name, value):
    self.set_data_property(name, "des", value)

  def set_data_property(self, name, propertyname="des", value=""):
    if name in self.__dict__:
      obj = self.__dict__[name]
      obj[propertyname] = value
      self.__dict__[name] = obj
    else:
      self.__dict__[name] = {propertyname: value}

  def set_data_property_countcell(self, name: str, value: int = 0):
    # if isproperty:
    # 普通求合列
    self.set_data_property(name, 'countcell', value)

  @staticmethod
  def get_countcell_num(dic):
    '''合计数值'''
    if isinstance(dic, dict) and 'countcell' in dic:
      return dic['countcell']
    return None

  @staticmethod
  def get_des_contain_countcell(dic):
    '''获取带合计列的列内容'''
    if isinstance(dic, dict) and 'des' in dic:
      return dic['des']
    return None