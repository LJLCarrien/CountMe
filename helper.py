from formatItem import FormatItem
from titleItem import TitleItem, TitleType
from configureData import ConfigureData
import xlsxwriter
from xlsxwriter.format import Format

# 英文简称获取：calendar.day_abbr
cn_day_abbr = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']

class XlsHelper():
    def __init__(self, configure: ConfigureData, excelSavePath: str):
        if configure is None:
            print('配置数据有误')
            return None
        self.configure = configure
        self.formatDic = configure.getFormatDic()
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

    def getTitleFormat(self, itemInfo: TitleItem, isSecTitle: bool = False) -> Format:
        '''行标题格式'''
        workbook = self.workbook
        type = TitleType.MAIN
        if isSecTitle:
            type = TitleType.SEC
        format = itemInfo.getFormat(workbook, self.configure.jsonDic, type)
        return format

    def getFormat(self, name: str) -> Format:
        if self.configure is None:
            print('配置数据有误')
            return None
        formatitem = FormatItem(self.configure.jsonDic, self.formatDic[name])
        format = formatitem.getFormat(self.workbook)
        return format

    def getDateFormat(self) -> Format:
        '''日期列格式'''
        date_format = self.getFormat('date')
        return date_format

    def getWeekFormat(self) -> Format:
        '''周目列格式'''
        weekday_format = self.getFormat('weekDay')
        return weekday_format

    def getCountNumFormat(self) -> Format:
        '''合计列内容格式'''
        countNumCell_format = self.getFormat('countNumCell')
        return countNumCell_format

    def getSumResultFormat(self) -> Format:
        '''求和内容格式'''
        sumResult_format = self.getFormat('sumResult')
        return sumResult_format

    def getMonthEndFormat(self) -> Format:
        '''月统计格式'''
        monthEnd_format = self.getFormat('monthEnd')
        return monthEnd_format


def getJsonData(jsonFilePath) -> ConfigureData:
    conf = ConfigureData(jsonFilePath)
    return conf
