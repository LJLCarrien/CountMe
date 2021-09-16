import json
from warnings import resetwarnings


class TitleItem(dict):
    def reset(self):
        self.name = ""
        self.width = None
        self.height = None
        self.fontName = None
        self.fontSize = None
        self.fontColor = None
        self.bgColor = None

    def __init__(self, info):
        self.reset()
        if isinstance(info, dict):
            for key, value in info.items():
                self.__dict__[key] = value


class ConfigureData():
    def __init__(self, jsonFilePath):
        self.filePath = jsonFilePath
        self.jsonDic = None
        self.defaultFontName = None
        self.titleFontSize = None
        self.secTitleFontSize = None
        self.titleFontColor = "black"
        self.titleBgColor = None
        self.secondTitle = {}
        self.defaultTitleWidth = None
        self.defaultTitleHeight = None
        self.openFile()

    def __listDicItem2class(self, keyName: str, tmpList: list) -> list:
        '''json数组里的dict对象转换成对应的类对象'''
        resultList = []
        for item in tmpList:
            if isinstance(item, dict):
                if keyName == 'title':
                    resultItem = TitleItem(item)
                    resultList.append(resultItem)
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
    def ifNoneInsteadDefault(defaultValue, value):
        if value is not None:
            return value
        return defaultValue

    def getSecTitle(self, itemInfo: TitleItem) -> list:
        '''获取一级菜单对应的二级菜单列表'''
        result = hasattr(itemInfo, 'secondSort')
        if result:
            key = itemInfo.secondSort
            if key in self.secondTitle:
                return self.secondTitle[key]
        return None

    def getTitleTotalLine(self) -> int:
        '''获取标题所占行数（包括二级标记在内）'''
        result = 1
        secNum = len(self.secondTitle)
        if secNum >= 1:
            result = result+1
        return result

    def getTitleWidth(self, itemInfo: TitleItem):
        '''列宽'''
        width = self.defaultTitleWidth
        if itemInfo.width is not None:
            width = itemInfo.width
        return width

    def getTitleHeight(self):
        '''行高'''
        return self.defaultTitleHeight

    def getCountCellValue(self, itemInfo: TitleItem):
        '''合计'''
        result = hasattr(itemInfo, 'countCell')
        if result:
            return itemInfo.countCell
        return None

    def keyTrueInTitleItem(self, key: str, itemInfo: TitleItem) -> bool:
        '''属性存在且为true，返回True'''
        isIn = hasattr(itemInfo, key)
        if isIn:
            isTrue = itemInfo.__getattribute__(key)
            return isIn and isTrue
        return False

    def getRowSumDicValue(self, itemInfo: TitleItem):
        '''行求和'''
        result = hasattr(itemInfo, 'rowSumDic')
        if result:
            return itemInfo.rowSumDic
        return None

    def getRowSumDicKeys(self) -> list:
        result = 'rowSumDicKey' in self.jsonDic
        if result:
            return self.rowSumDicKey
        return None

    def getColSumDicValue(self, itemInfo: TitleItem):
        '''列求和'''
        result = hasattr(itemInfo, 'colSumDic')
        if result:
            return itemInfo.colSumDic
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


def getJsonData(jsonFilePath) -> ConfigureData:
    conf = ConfigureData(jsonFilePath)
    return conf