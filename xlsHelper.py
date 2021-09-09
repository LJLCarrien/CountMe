import xlsxwriter
from xlsxwriter.format import Format
from jsonHelper import ConfigureData, TitleItem


class xlsHelper():
    def __init__(self, configure: ConfigureData, excelSavePath: str):
        self.configure = configure
        self.workbook = xlsxwriter.Workbook(excelSavePath)

    def close_workbook(self):
        return self.workbook.close

    def get_worksheet(self, sheetName: str):
        index = self.workbook._get_sheet_index(sheetName)
        if index is None:
            self.worksheet = self.workbook.add_worksheet(sheetName)
        else:
            self.worksheet = self.workbook.get_worksheet_by_name(sheetName)
        return self.worksheet

    def getNewTitleFormat(self, itemInfo: TitleItem, isSecTitle=False) -> Format:
        if self.configure is None:
            print('配置数据有误')
            return
        conf = self.configure

        format = self.workbook.add_format()
        format.set_bold()
        format.set_align('center')
        format.set_align('vcenter')

        # fontSize
        if isSecTitle:
            fontSize = ConfigureData.ifNoneInsteadDefault(
                conf.secTitleFontSize, itemInfo.fontSize)
        else:
            fontSize = ConfigureData.ifNoneInsteadDefault(
                conf.titleFontSize, itemInfo.fontSize)
        format.set_font_size(fontSize)

        # fontName
        fontName = ConfigureData.ifNoneInsteadDefault(
            conf.titleFontName, itemInfo.fontName)
        format.set_font_name(fontName)

        # fontColor
        fontColor = ConfigureData.ifNoneInsteadDefault(
            conf.titleFontColor, itemInfo.fontColor)
        format.set_font_color(fontColor)

        # bgColor
        bgColor = ConfigureData.ifNoneInsteadDefault(
            conf.titleBgColor, itemInfo.bgColor)
        format.set_bg_color(bgColor)
        return format
