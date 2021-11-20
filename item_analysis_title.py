from enum import Enum

import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.workbook import Workbook

from item_format import FormatItem


class AnalysisTitleType(Enum):
  '''结果分析-标题类型'''
  COL = 1
  '''列标题'''
  ROW = 2
  '''行标题'''


class AnalysisTitleItem():
  '''结果分析-标题'''
  def __init__(self) -> None:
    self.name = ""
    self.showname = ""
    self.bold = None
    self.fontsize = None
    self.fontname = None
    self.fontcolor = None
    self.bgcolor = None
    '''item:行标题，对应该行的其他格子;列标题，对应该列的其他格子'''
    self.item_fontcolor = None
    self.item_bgcolor = None
    self.type = None
    self.formatitem: FormatItem = None
    self.format: Format = None
    self.cellformat: Format = None
    self.contenttype = None

  def set_data(self, info: dict):
    for key, value in info.items():
      if key == 'type':
        print("{key} 为 AnalysisTitleItem 保留默认格式属性字段,不支持在 analysis_coltitle 配置")
      else:
        self.__dict__[key] = value
        # if not hasattr(self, key):
        #     print("TitleItem 格式属性：%s未支持" % key)

  def get_defaultformat_by_type(self, confdic: dict) -> FormatItem:
    '''获取行标题、列标题的默认格式'''
    if self.type == AnalysisTitleType.COL:
      f_item = FormatItem(confdic, confdic['format']['analysis_col_title'])
    elif self.type == AnalysisTitleType.ROW:
      f_item = FormatItem(confdic, confdic['format']['analysis_row_title'])
    return f_item

  def get_format(self, workbook: Workbook, confdic: dict) -> Format:
    if self.format is not None:
      return self.format
    f_item = self.get_defaultformat_by_type(confdic)
    self.formatitem = f_item
    self.formatitem.item_fontcolor = self.item_fontcolor
    self.formatitem.item_bgcolor = self.item_bgcolor
    format = self.formatitem.get_format(workbook)
    if self.bold is not None:
      format.set_bold(self.bold)
    if self.fontsize is not None:
      format.set_font_size(self.fontsize)
    if self.fontname is not None:
      format.set_font_name(self.fontname)
    if self.fontcolor is not None:
      format.set_font_color(self.fontcolor)
    if self.bgcolor is not None:
      format.set_bg_color(self.bgcolor)
    self.format = format
    return format

  def get_itemcell_format(self, workbook: Workbook) -> Format:
    '''获取行/列标题的列格子内容格式'''
    # if self.cellformat is not None:
    #   return self.cellformat
    format = self.formatitem.get_itemcell_format(workbook)
    self.cellformat = format
    return format

  def get_isneed_handle_itemcell(self):
    return self.item_bgcolor is not None or self.item_fontcolor is not None

  def get_is_data(self):
    '''分析表格的输入数据'''
    return self.contenttype == "data"


class ColAnalysisTitleItem(AnalysisTitleItem):
  '''列标题'''
  def __init__(self) -> None:
    super().__init__()
    self.type = AnalysisTitleType.COL


class RowAnalysisTitleItem(AnalysisTitleItem):
  '''行标题'''
  def __init__(self) -> None:
    super().__init__()
    self.format_cover_col = True
    self.type = AnalysisTitleType.ROW
