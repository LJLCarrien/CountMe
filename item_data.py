class DataItem():
  def __init__(self, date="") -> None:
    self.date = date
    self.year = 0
    self.month = 0
    self.day = 0

  @property
  def date(self):
    return self.__date

  @date.setter
  def date(self, value: str):
    self.__date = value
    if value == "":
      return
    datalist = value.split('/')
    self.year = int(datalist[0])
    self.month = int(datalist[1])
    self.day = int(datalist[2])

  @property
  def row(self):
    return self.__row

  @row.setter
  def row(self, value):
    self.__row = value

  def set_data(self, name, value=""):
    self.__dict__[name] = value

  def get_data(self, name):
    return self.__dict__[name]

  def set_des(self, name, value):
    self.set_data_property(name, "des", value)

  def set_data_property(self, name, propertyname="des", value=""):
    if name in self.__dict__:
      obj = self.__dict__[name]
      obj[propertyname] = value
      self.__dict__[name] = obj
    else:
      self.__dict__[name] = {propertyname: value}

  def get_data_property(self, name, propertyname):
    if name in self.__dict__:
      if propertyname in self.__dict__[name]:
        return self.__dict__[name][propertyname]
    return None

  def cal_countcell(self, name, col_name_list: list):
    result = 0
    for colname in col_name_list:
      for key, value in self.__dict__.items():
        if key == colname:
          num = self.get_countcell_num(value)
          if num is not None:
            result = result + num
            break
    self.set_data_property_countcell(name, result)

  def set_data_property_countcell(self, name: str, value: int = 0):
    '''设置合计数值'''
    self.set_data_property(name, 'countcell', value)

  def set_weekday_index(self, value: int):
    '''设置周目下标'''
    self.set_data('weekday', value)

  def get_weekday_index(self) -> int:
    '''周目下标'''
    return self.get_data('weekday')

  def get_daycoust_countcell(self):
    '''获取日合计值'''
    result = self.get_data_property('daycoust', 'countcell')
    return result

  @staticmethod
  def get_countcell_num(dic):
    '''获取合计数值'''
    if isinstance(dic, dict) and 'countcell' in dic:
      return dic['countcell']
    return None

  @staticmethod
  def get_des_contain_countcell(dic):
    '''获取带合计列的列内容'''
    if isinstance(dic, dict) and 'des' in dic:
      return dic['des']
    return None