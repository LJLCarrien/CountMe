import json
from typing import Any
import xlsxwriter
from xlsxwriter.format import Format
from enum import Enum

# 英文简称获取：calendar.day_abbr
cn_day_abbr = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']


class TitleType(Enum):
    '''主标题'''
    MAIN = 0
    '''二级标题'''
    SEC = 1


class ConfigureData():
    def __init__(self, jsonFilePath):
        self.filePath = jsonFilePath
        self.jsonDic = None
        self.secondTitle = {}

        self.defaultTitleWidth = None
        self.defaultTitleHeight = None

        self.defaultFontName = None
        self.defaultFontSize = None
        self.defaultFontColor = None
        self.defaultBgColor = None
        self.defaultHorAlignment = None
        self.defaultVerAlignment = None

        self.openFile()

    def __listDicItem2class(self, keyName: str, tmpList: list) -> list:
        '''json数组里的dict对象转换成对应的类对象'''
        resultList = []
        for item in tmpList:
            if isinstance(item, dict):
                if keyName == 'title':
                    titleitem = TitleItem()
                    titleitem.setData(self.jsonDic, item)
                    resultList.append(titleitem)
                else:
                    print('%s未支持' % keyName)
            else:
                return tmpList
        return resultList

    def openFile(self):
        if self.filePath == None:
            return None
        with open(self.filePath, 'rb')as f:
            jsonDic = json.load(f)
            self.jsonDic = jsonDic
            if isinstance(jsonDic, dict):
                for key, value in jsonDic.items():
                    if isinstance(value, list):
                        self.__dict__[key] = self.__listDicItem2class(
                            key, value)
                    else:
                        self.__dict__[key] = value

    def getTitleList(self) -> list:
        '''获取一级菜单列表'''
        return self.title

    @staticmethod
    def ifNoneUseDefault(defaultValue, value):
        if value is not None:
            return value
        return defaultValue

    def getTitleTotalLine(self) -> int:
        '''获取标题所占行数（包括二级标记在内）'''
        result = 1
        secNum = len(self.secondTitle)
        if secNum >= 1:
            result = result+1
        return result

    def getTitleHeight(self):
        '''行高'''
        return self.defaultTitleHeight

    def getRowSumDicKeys(self) -> list:
        result = 'rowSumDicKey' in self.jsonDic
        if result:
            return self.rowSumDicKey
        return None

    def getColSumDicKeys(self) -> list:
        result = 'colSumDicKey' in self.jsonDic
        if result:
            return self.colSumDicKey
        return None

    def getWeekCountKey(self) -> str:
        '''周合计'''
        return "weekCount"

    def getMonthCountKey(self) -> str:
        '''月合计'''
        return "monthCount"

    def getFormatDic(self) -> dict:
        '''获取fomat'''
        result = 'format' in self.jsonDic
        if result:
            return self.format
        return None


class FormatItem():

    def __init__(self, jsonDic: dict, info: Any = None) -> None:
        if jsonDic is None:
            print('配置数据有误')
            return
        self.fontSize = jsonDic["defaultFontSize"]
        self.bold = jsonDic["defaultBold"]
        self.fontName = jsonDic["defaultFontName"]
        self.horAlignment = jsonDic["defaultHorAlignment"]
        self.verAlignment = jsonDic["defaultVerAlignment"]
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

        if self.bold:
            format.set_bold()
        if self.numFormat is not None:
            format.set_num_format(self.numFormat)
        format.set_font_name(self.fontName)
        format.set_font_size(self.fontSize)
        format.set_align(self.horAlignment)
        format.set_align(self.verAlignment)
        if self.fontColor is not None:
            format.set_font_color(self.fontColor)
        if self.bgColor is not None:
            format.set_bg_color(self.bgColor)

        return format


class TitleItem(dict):
    def reset(self):
        self.name = ""
        self.secList = None
        # 标题默认格式
        self.formatItem: FormatItem = None
        # 标题定制格式
        self.fontName = None
        self.fontSize = None
        self.fontColor = None
        self.bgColor = None
        self.width = None
        self.height = None

    def __init__(self):
        self.reset()

    def setData(self, confDic: dict, info: dict):
        # 定制格式
        for key, value in info.items():
            self.__dict__[key] = value
            # if not hasattr(self, key):
            #     print("TitleItem 格式属性：%s未支持" % key)
        self.setSecList(confDic)
        self.setTitleWidth(confDic)

    def getTitleDefaultFormatByType(self, confDic: dict, type: TitleType) -> FormatItem:
        '''获取标题的默认格式'''
        if type == TitleType.MAIN:
            fItem = FormatItem(confDic, confDic['format']['title'])
        elif type == TitleType.SEC:
            fItem = FormatItem(confDic, confDic['format']['secTitle'])
        return fItem

    def getFormat(self, workbook: xlsxwriter.Workbook, confDic: dict, type: TitleType = TitleType.MAIN) -> Format:
        # 默认格式
        fItem = self.getTitleDefaultFormatByType(confDic, type)
        self.formatItem = fItem
        format = self.formatItem.getFormat(workbook)
        if self.fontName is not None:
            format.set_font_name(self.fontName)
        if self.fontSize is not None:
            format.set_font_size(self.fontSize)
        if self.fontColor is not None:
            format.set_font_color(self.fontColor)
        if self.bgColor is not None:
            format.set_bg_color(self.bgColor)
        return format

    def setSecList(self, confDic: dict):
        '''获取一级菜单对应的二级菜单列表'''
        result = hasattr(self, 'secondSort')
        if result:
            key = self.secondSort
            if key in confDic["secondTitle"]:
                self.secList = confDic["secondTitle"][key]

    def getSecList(self):
        return self.secList

    def setTitleWidth(self, confDic: dict):
        '''列宽'''
        width = confDic["defaultTitleWidth"]
        if self.width is not None:
            width = self.width
        self.width = width

    def getTitleWidth(self):
        return self.width

    def getCountCellValue(self):
        '''合计'''
        result = hasattr(self, 'countCell')
        if result:
            return self.countCell
        return None

    def getRowSumDicValue(self):
        '''行求和'''
        result = hasattr(self, 'rowSumDic')
        if result:
            return self.rowSumDic
        return None

    def getColSumDicValue(self):
        '''列求和'''
        result = hasattr(self, 'colSumDic')
        if result:
            return self.colSumDic
        return None

    def keyTrueInTitleItem(self, key: str) -> bool:
        '''属性存在且为true，返回True'''
        isIn = hasattr(self, key)
        if isIn:
            isTrue = self.__getattribute__(key)
            return isIn and isTrue
        return False


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
