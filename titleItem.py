from enum import Enum
import xlsxwriter

from xlsxwriter.format import Format
from formatItem import FormatItem

class TitleType(Enum):
    '''主标题'''
    MAIN = 0
    '''二级标题'''
    SEC = 1

class TitleItem(dict):
    def reset(self):
        self.name = ""
        self.secList = None
        # 标题默认格式
        self.formatItem: FormatItem = None
        # 标题定制格式
        self.fontSize = None
        self.fontName = None
        self.fontColor = None
        self.bgColor = None
        self.width = None
        self.height = None

    def __init__(self):
        self.reset()

    def setData(self, confDic: dict, info: dict):
        # 定制格式
        for key, value in info.items():
            if key=='formatItem':
                print("%s 为 TitleItem 保留默认格式属性字段,不支持在 title 配置" % key)
            elif key=='secList':
                print("%s 为 TitleItem 保留默认格式属性字段,不支持在 secList 配置" % key)
            else:
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
        if self.fontSize is not None:
            format.set_font_size(self.fontSize)
        if self.fontName is not None:
            format.set_font_name(self.fontName)
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
