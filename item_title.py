from enum import Enum
import xlsxwriter

from xlsxwriter.format import Format
from item_format import FormatItem


class TitleType(Enum):
  """标题类型"""
  MAIN = 0
  '''主标题'''
  SEC = 1
  '''二级标题'''


class TitleItem():
  """标题类"""
  def reset(self):
    self.name = ""
    self.showname = ""
    self.sec_list = None
    # 标题默认格式
    self.formatitem: FormatItem = None
    # 标题定制格式
    self.fontsize = None
    self.fontname = None
    self.fontcolor = None
    self.bgcolor = None
    self.width = None
    self.height = None
    # 列下标
    self.columnindex = []

  def __init__(self):
    self.reset()

  def set_columindex(self, index):
    # 设置列下标
    self.columnindex.append(index)

  def get_countcell_col_index(self):
    '''获取合计列下标'''
    bhave_countcell = self.get_addcountcell_value() is not None
    bhave_seclist = len(self.columnindex) > 1
    if bhave_seclist:
      if bhave_countcell:
        # 二级列标题&有合计,最后一个元素列+1，如餐饮
        return self.columnindex[-1] + 1
      else:
        print('目前未有该逻辑：二级列表，没有合计列')
        return 0
    else:
      if bhave_countcell:
        # 一级列标题，有合计列，如生活用品、交通等
        return self.columnindex[0] + 1
      else:
        # 一级列标题，内容即合计，如日必须，日合计
        return self.columnindex[0]

  def set_data(self, confdic: dict, info: dict):
    # 定制格式
    for key, value in info.items():
      if key == 'formatitem':
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
    self.formatitem = f_item
    format = self.formatitem.get_format(workbook)
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

  def get_is_seclist(self) -> bool:
    '''是否有二级列表'''
    if self.sec_list is None:
      return False
    return len(self.sec_list) > 0

  def get_seclist_name_by_index(self, index):
    return self.sec_list[index]['name']

  def get_seclist_showname_by_index(self, index):
    return self.sec_list[index]['showname']

  def set_title_width(self, confdic: dict):
    '''列宽'''
    width = confdic["defaultTitleWidth"]
    if self.width is not None:
      width = self.width
    self.width = width

  def get_title_width(self):
    return self.width

  def get_addcountcell_value(self):
    '''增加合计列标题'''
    result = hasattr(self, 'addCountCell')
    if result:
      return self.addCountCell
    return None

  def get_countcell(self) -> bool:
    '''是否用于专门用于求和的数字列'''
    result = hasattr(self, 'countcell')
    if result:
      return self.countcell
    return False

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
