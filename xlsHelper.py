import xlsxwriter
from xlsxwriter.format import Format
from handleConfig import ConfigureData, TitleItem

cn_day_abbr = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']


class xlsHelper():
    def __init__(self, configure: ConfigureData, excelSavePath: str):
        # 英文简称获取：calendar.day_abbr
        self.configure = configure
        self.workbook = xlsxwriter.Workbook(excelSavePath)

    def close_workbook(self):
        return self.workbook.close()

    @staticmethod
    def getWeekDayCnNameByIndex(weekDayIndex: int) -> str:
        return cn_day_abbr[weekDayIndex]

    @staticmethod
    def rowCol2Str(row, col):
        '''行列转字符串
        eq: 0,0->A1 ; 1,0->A2 ; 0，1->B1 ;
        '''
        result = "%s%s" % (chr(ord('A')+col), row+1)
        return result

    def get_worksheet(self, sheetName: str):
        index = self.workbook._get_sheet_index(sheetName)
        if index is None:
            self.worksheet = self.workbook.add_worksheet(sheetName)
        else:
            self.worksheet = self.workbook.get_worksheet_by_name(sheetName)
        return self.worksheet

    def getNewTitleFormat(self, itemInfo: TitleItem, isSecTitle=False) -> Format:
        '''行标题格式'''
        if self.configure is None:
            print('配置数据有误')
            return None
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
            conf.defaultFontName, itemInfo.fontName)
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

    def getDateFormat(self) -> Format:
        '''日期列格式'''
        if self.configure is None:
            print('配置数据有误')
            return None
        date_format = self.workbook.add_format({'num_format': 'm"月"d"日"'})
        date_format.set_align('center')
        date_format.set_align('vcenter')
        # fontName
        fontName = self.configure.defaultFontName
        date_format.set_font_name(fontName)
        return date_format

    def getWeekFormat(self) -> Format:
        '''周目列格式'''
        if self.configure is None:
            print('配置数据有误')
            return None
        weekday_format = self.workbook.add_format()
        weekday_format.set_align('left')
        weekday_format.set_align('vcenter')
        # fontName
        fontName = self.configure.defaultFontName
        weekday_format.set_font_name(fontName)
        return weekday_format

    def getCountNumFormat(self) -> Format:
        '''合计列内容格式'''
        countNumCell_format = self.workbook.add_format({'num_format': '0.00_'})
        countNumCell_format.set_align('center')
        countNumCell_format.set_align('vcenter')
        countNumCell_format.set_font_size(11)
        countNumCell_format.set_font_name('宋体')
        countNumCell_format.set_font_color('red')

    def getSumResultFormat(self) -> Format:
        '''求和内容格式'''
        sumResult_format = self.workbook.add_format()
        sumResult_format.set_bold()
        sumResult_format.set_align('center')
        sumResult_format.set_align('vcenter')
        sumResult_format.set_font_size(16)
        sumResult_format.set_font_name('宋体')
        sumResult_format.set_font_color('red')
        return sumResult_format

    def getMonthEndFormat(self) -> Format:
        monthEnd_format = self.workbook.add_format()
        monthEnd_format.set_bold()
        monthEnd_format.set_align('center')
        monthEnd_format.set_align('vcenter')
        monthEnd_format.set_font_size(16)
        monthEnd_format.set_font_name('宋体')
        monthEnd_format.set_bg_color('#dca09e')
        monthEnd_format.set_font_color('#a12f2f')
        return monthEnd_format
