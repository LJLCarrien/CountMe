import json
from titleItem import TitleItem

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
    
    def getExcelPath(self)->str:
        '''获取保存'''
        result = 'excelPath' in self.jsonDic
        if result:
            return self.excelPath
        return None
