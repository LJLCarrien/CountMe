from typing import Any

import xlsxwriter
from xlsxwriter.format import Format


class FormatItem():
  """格式类"""
  def __init__(self, jsonDic: dict, info: Any = None) -> None:
    if jsonDic is None:
      print('配置数据有误')
      return
    self.bold = jsonDic["defaultBold"]
    self.horalignment = jsonDic["defaultHorAlignment"]
    self.veralignment = jsonDic["defaultVerAlignment"]
    self.fontsize = jsonDic["defaultFontSize"]
    self.fontname = jsonDic["defaultFontName"]
    self.fontcolor = jsonDic["defaultFontColor"]
    self.bgcolor = jsonDic["defaultBgColor"]
    self.numformat = None
    if info is not None:
      self.set_data(info)

  def set_data(self, info: Any):
    if isinstance(info, dict):
      dic = info
      for key in dic:
        if hasattr(self, key):
          self.__setattr__(key, dic[key])
        else:
          print(f"item_format 格式属性：{key}未支持")
    else:
      print(f"item_format setData 未支持该类型: {type(info)}")

  def get_format(self, workbook: xlsxwriter.Workbook) -> Format:
    result_format = workbook.add_format()
    if self.numformat is not None:
      result_format.set_num_format(self.numformat)
    if self.bold:
      result_format.set_bold()
    result_format.set_align(self.horalignment)
    result_format.set_align(self.veralignment)
    result_format.set_font_size(self.fontsize)
    result_format.set_font_name(self.fontname)
    if self.fontcolor is not None:
      result_format.set_font_color(self.fontcolor)
    if self.bgcolor is not None:
      result_format.set_bg_color(self.bgcolor)
    return result_format
