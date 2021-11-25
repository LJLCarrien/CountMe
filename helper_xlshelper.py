from xlsxwriter import worksheet
from xlsxwriter.chart import Chart
from configure_data import ConfigureData
from item_analysis_title import AnalysisTitleItem
from item_format import FormatItem
from item_title import TitleItem, TitleType
import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol

# 英文简称获取：calendar.day_abbr
cn_day_abbr = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']


class XlsHelper():
  """excel帮助类"""
  def __init__(self, configure: ConfigureData, excel_savepath: str):
    if configure is None:
      print('配置数据有误')
      return None
    self.configure = configure
    self.formatdic = configure.get_format_dic()
    self.workbook = xlsxwriter.Workbook(excel_savepath)

  def close_workbook(self):
    return self.workbook.close()

  @staticmethod
  def get_weekday_cn_name_by_index(weekday_index: int) -> str:
    return cn_day_abbr[weekday_index]

  @staticmethod
  def get_rowcol_2_str(row, col):
    '''行列转字符串 
    eq: 0,0->A1 ; 1,0->A2 ; 0，1->B1 ;
    '''
    # result = f"{chr(ord('A') + col)}{row + 1}"
    result = xl_rowcol_to_cell(row, col)
    return result

  @staticmethod
  def get_str_2_rowcol(str):
    '''字符串转行列
    xl_cell_to_rowcol('A1')    # (0, 0)
    '''
    (row, col) = xl_cell_to_rowcol(str)
    return (row, col)

  def get_worksheet(self, sheetname: str):
    index = self.workbook._get_sheet_index(sheetname)
    if index is None:
      self.worksheet = self.workbook.add_worksheet(sheetname)
    else:
      self.worksheet = self.workbook.get_worksheet_by_name(sheetname)
    return self.worksheet

  def get_chart(self, typename: str) -> Chart:
    return self.workbook.add_chart({'type': typename})

  def get_titleformat(self,
                      iteminfo: TitleItem,
                      is_sectitle: bool = False) -> Format:
    '''行标题格式'''
    workbook = self.workbook
    type = TitleType.MAIN
    if is_sectitle:
      type = TitleType.SEC
    format = iteminfo.get_format(workbook, self.configure.jsondic, type)
    return format

  def get_format(self, name: str) -> Format:
    if self.configure is None:
      print('配置数据有误')
      return None
    formatitem = FormatItem(self.configure.jsondic, self.formatdic[name])
    format = formatitem.get_format(self.workbook)
    return format

  def get_dateformat(self) -> Format:
    '''日期列格式'''
    date_format = self.get_format('date')
    return date_format

  def get_weekformat(self) -> Format:
    '''周目列格式'''
    weekday_format = self.get_format('weekDay')
    return weekday_format

  def get_countnum_format(self) -> Format:
    '''合计列内容格式'''
    countnumcell_format = self.get_format('countNumCell')
    return countnumcell_format

  def get_sumresult_format(self) -> Format:
    '''求和内容格式'''
    sumresult_format = self.get_format('sumResult')
    return sumresult_format

  def get_monthend_format(self) -> Format:
    '''月统计格式'''
    monthend_format = self.get_format('monthEnd')
    return monthend_format

  def get_analysis_titleitem_format(self,
                                    iteminfo: AnalysisTitleItem) -> Format:
    '''分析结果-行/列标题格式'''
    workbook = self.workbook
    format = iteminfo.get_format(workbook, self.configure.jsondic)
    return format

  def get_analysis_itemcell_format(self,
                                   iteminfo: AnalysisTitleItem) -> Format:
    '''分析结果-行/列标题对应行/列内容格式'''
    workbook = self.workbook
    format = iteminfo.get_itemcell_format(workbook)
    return format


def get_jsondata(json_file_path) -> ConfigureData:
  conf = ConfigureData(json_file_path)
  return conf
