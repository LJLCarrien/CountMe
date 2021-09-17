from typing import Any

import xlsxwriter
from xlsxwriter.format import Format

class FormatItem():

    def __init__(self, jsonDic: dict, info: Any = None) -> None:
        if jsonDic is None:
            print('配置数据有误')
            return
        self.bold = jsonDic["defaultBold"]
        self.horAlignment = jsonDic["defaultHorAlignment"]
        self.verAlignment = jsonDic["defaultVerAlignment"]
        self.fontSize = jsonDic["defaultFontSize"]
        self.fontName = jsonDic["defaultFontName"]
        self.fontColor = jsonDic["defaultFontColor"]
        self.bgColor = jsonDic["defaultBgColor"]
        self.numFormat = None
        if info is not None:
            self.setData(info)

    def setData(self, info: Any):
        if isinstance(info, dict):
            dic = info
            for key in dic:
                if hasattr(self, key):
                    self.__setattr__(key, dic[key])
                else:
                    print("formatItem 格式属性：%s未支持" % key)
        else:
            print('formatItem setData 未支持该类型: %s' % type(info))

    def getFormat(self, workbook: xlsxwriter.Workbook) -> Format:
        format = workbook.add_format()

        if self.numFormat is not None:
            format.set_num_format(self.numFormat)
        if self.bold:
            format.set_bold()
        format.set_align(self.horAlignment)
        format.set_align(self.verAlignment)
        format.set_font_size(self.fontSize)
        format.set_font_name(self.fontName)
        if self.fontColor is not None:
            format.set_font_color(self.fontColor)
        if self.bgColor is not None:
            format.set_bg_color(self.bgColor)
        return format