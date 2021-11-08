from enum import Enum
import xlsxwriter

from xlsxwriter.format import Format
from format_item import FormatItem


class TitleType(Enum):
  """标题类型"""
  '''主标题'''
  MAIN = 0
  '''二级标题'''
  SEC = 1


class TitleItem():
  """标题类"""
  def reset(self):
    self.name = ""
    self.sec_list = None
    # 标题默认格式
    self.format_item: FormatItem = None
    # 标题定制格式
    self.fontsize = None
    self.fontname = None
    self.fontcolor = None
    self.bgcolor = None
    self.width = None
    self.height = None

  def __init__(self):
    self.reset()

  def set_data(self, confdic: dict, info: dict):
    # 定制格式
    for key, value in info.items():
      if key == 'format_item':
        print("{key} 为 TitleItem 保留默认格式属性字段,不支持在 title 配置")
      elif key == 'secList':
        print("{key} 为 TitleItem 保留默认格式属性字段,不支持在 secList 配置")
      else:
        self.__dict__[key] = value
        # if not hasattr(self, key):
        #     print("TitleItem 格式属性：%s未支持" % key)
    self.set_seclist(confdic)
    self.set_title_width(confdic)

  def get_title_defaultformat_by_type(self, confdic: dict,
                                      type: TitleType) -> FormatItem:
    '''获取标题的默认格式'''
    if type == TitleType.MAIN:
      f_item = FormatItem(confdic, confdic['format']['title'])
    elif type == TitleType.SEC:
      f_item = FormatItem(confdic, confdic['format']['secTitle'])
    return f_item

  def get_format(self,
                 workbook: xlsxwriter.Workbook,
                 confdic: dict,
                 type: TitleType = TitleType.MAIN) -> Format:
    # 默认格式
    f_item = self.get_title_defaultformat_by_type(confdic, type)
    self.format_item = f_item
    format = self.format_item.get_format(workbook)
    if self.fontsize is not None:
      format.set_font_size(self.fontsize)
    if self.fontname is not None:
      format.set_font_name(self.fontname)
    if self.fontcolor is not None:
      format.set_font_color(self.fontcolor)
    if self.bgcolor is not None:
      format.set_bg_color(self.bgcolor)
    return format

  def set_seclist(self, confdic: dict):
    '''获取一级菜单对应的二级菜单列表'''
    result = hasattr(self, 'secondSort')
    if result:
      key = self.secondSort
      if key in confdic["secondTitle"]:
        self.sec_list = confdic["secondTitle"][key]

  def get_seclist(self):
    return self.sec_list

  def set_title_width(self, confdic: dict):
    '''列宽'''
    width = confdic["defaultTitleWidth"]
    if self.width is not None:
      width = self.width
    self.width = width

  def get_title_width(self):
    return self.width

  def get_countcell_value(self):
    '''合计'''
    result = hasattr(self, 'countCell')
    if result:
      return self.countCell
    return None

  def get_rowsum_dicvalue(self):
    '''行求和'''
    result = hasattr(self, 'rowSumDic')
    if result:
      return self.rowSumDic
    return None

  def get_colsum_dicvalue(self):
    '''列求和'''
    result = hasattr(self, 'colSumDic')
    if result:
      return self.colSumDic
    return None

  def get_isfreeze(self):
    '''冻结'''
    result = hasattr(self, 'freeze')
    if result:
      return self.freeze
    return False

  def get_is_keytrue_in_titleitem(self, key: str) -> bool:
    '''属性存在且为true，返回True'''
    isin = hasattr(self, key)
    if isin:
      istrue = self.__getattribute__(key)
      return isin and istrue
    return False
