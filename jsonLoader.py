import json


class TitleItem():
    def reset(self):
        self.name = ""
        self.fontName = "宋体"
        self.fontSize = 10
        self.fontColor = "white"
        self.bgColor = ""
        self.width = "10"
        self.height = "10"

    def __init__(self, info):
        if isinstance(info, dict):
            for key, value in info.items():
                self[key] = value
        pass


class ConfigureData():
    def __init__(self, jsonFilePath):
        self.filePath = jsonFilePath
        self.Info = {}
        self.openFile()

    def openFile(self):
        if self.filePath == None:
            return None
        with open(self.filePath, 'rb')as f:
            jsonDic = json.load(f)
            print(jsonDic)
        if isinstance(jsonDic, dict):
            for key, value in jsonDic.items():
                self.Info[key] = value
                if key == "title" and isinstance(value, list):
                    titleList = []
                    for item in value:
                        print(item)
                        # titleList.append(TitleItem(item))

            print(self.Info)


class Dict(dict):
    __setattr__ = dict.__setitem__
    __getattr__ = dict.__getitem__


def dictToObj(dictObj):
    if not isinstance(dictObj, dict):
        return dictObj
    d = Dict()
    for k, v in dictObj.items():
        d[k] = dictToObj(v)
    return d


def getJsonData(jsonFilePath):
    conf = ConfigureData(jsonFilePath)
