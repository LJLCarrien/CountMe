from datetime import datetime
import calendar
from helper import TitleItem, XlsHelper, getJsonData
import operator

excelSavePath = './test.xlsx'
obj = calendar.Calendar()

c_year = datetime.now().year
max_month = 12

rowSumDic = {}
colSumDic = {}
# 合计列下标数组
countCellList = []

jsonFilePath = './config.json'
jsonInfo = getJsonData(jsonFilePath)


class RowSumResult(object):
    def __init__(self):
        # 求和结果-列下标
        self.index = -1
        # 用于求和-列下标
        self.itemList = []

    def setIndex(self, i):
        '''求和结果-列下标'''
        self.index = i

    def setList(self, list):
        self.itemList = list

    def addListItem(self, item):
        '''求和元素-列下标'''
        if self.itemList == None:
            self.itemList = []
        self.itemList.append(item)


def rowSumSetResultIndex(key, content):
    '''
    行求和结果，列下标设置
    '''
    if key not in rowSumDic:
        rowSumDic[key] = RowSumResult()
    tmp = rowSumDic[key]
    if isinstance(tmp, RowSumResult):
        tmp.setIndex(content)


def rowSumAddItemIndex(key, listItem):
    '''
    行求和元素，列下标设置
    '''
    if key not in rowSumDic:
        rowSumDic[key] = RowSumResult()
    tmp = rowSumDic[key]
    if isinstance(tmp, RowSumResult):
        tmp.addListItem(listItem)


class ColSumResult(object):
    '''
    列求和
    '''

    def __init__(self):
        # 求和结果-列下标
        self.resultIndex = -1
        # 求和对象-列下标
        self.OperatorIndex = -1
        # 求和对象-行开始
        self.rowStartIndex = -1
        # 求和对象-行结束
        self.rowEndIndex = -1

    def setResultIndex(self, i):
        self.resultIndex = i

    def setOperatorIndex(self, i):
        self.OperatorIndex = i

    def setRowStartIndex(self, i):
        self.rowStartIndex = i

    def setRowEndIndex(self, i):
        self.rowEndIndex = i

    def getSumStr(self, startIndex, endIndex):
        self.setRowStartIndex(startIndex)
        self.setRowEndIndex(endIndex)
        if self.OperatorIndex != -1:
            beginStr = XlsHelper.rowCol2Str(self.rowStartIndex, self.OperatorIndex)
            endStr = XlsHelper.rowCol2Str(self.rowEndIndex, self.OperatorIndex)
            return "=SUM(%s:%s)" % (beginStr, endStr)
        return ""


def colSumDicCallFunc(key, funName, content):
    '''
    列求和字典,列下标设置
    '''
    if key not in colSumDic:
        colSumDic[key] = ColSumResult()
    mc = operator.methodcaller(funName, content)
    mc(colSumDic[key])


def getSumStr(key, startIndex, endIndex):
    if key not in colSumDic:
        colSumDic[key] = ColSumResult()
    tmp = colSumDic[key]
    if isinstance(tmp, ColSumResult):
        return tmp.getSumStr(startIndex, endIndex)
    return None


def main():
    helper = XlsHelper(jsonInfo, excelSavePath)
    worksheet = helper.get_worksheet(str(c_year))
    titleList = jsonInfo.getTitleList()
    rowSumDicKey = jsonInfo.getRowSumDicKeys()
    colSumDicKey = jsonInfo.getColSumDicKeys()

    lineIndex = 0  # 行下标
    listIndex = 0  # 列下标
    for titleItem in titleList:
        if isinstance(titleItem, TitleItem):
            value = titleItem.name
            titleFormat = helper.getTitleFormat(titleItem)

            secondMenuList = titleItem.getSecList()
            isHaveSecTitle = secondMenuList is not None
            if isHaveSecTitle:
                oldlistIndex = listIndex
                secondMenuLen = len(secondMenuList)
                listIndex = listIndex+secondMenuLen-1
                # 一级菜单合并-行列行列
                worksheet.merge_range(lineIndex, oldlistIndex, lineIndex, listIndex, value, titleFormat)
                # 二级菜单内容
                SecTitleFormat = helper.getTitleFormat(titleItem, True)
                for tmpLie in range(secondMenuLen):
                    worksheet.write(lineIndex+1, oldlistIndex + tmpLie, secondMenuList[tmpLie], SecTitleFormat)
            else:
                worksheet.merge_range(lineIndex, listIndex, lineIndex+1, listIndex, value, titleFormat)
            # 列宽行高
            titleWidth = titleItem.getTitleWidth()
            worksheet.set_column(listIndex, listIndex, titleWidth)
            titleHeight = jsonInfo.getTitleHeight()
            worksheet.set_row(lineIndex, titleHeight)

            # 合计
            countCellValue = titleItem.getCountCellValue()
            isHaveCountCell = countCellValue is not None
            if isHaveCountCell:
                listIndex = listIndex+1
                countCellList.append(listIndex)
                worksheet.merge_range(lineIndex, listIndex, lineIndex+1, listIndex, countCellValue, titleFormat)

            # 行求和
            for key in rowSumDicKey:
                isNeedRowSum = titleItem.keyTrueInTitleItem(key)
                if isNeedRowSum:
                    rowSumAddItemIndex(key, listIndex)

            rowSumDicValue = titleItem.getRowSumDicValue()
            isRowSumResultHere = rowSumDicValue is not None
            if isRowSumResultHere:
                rowSumSetResultIndex(rowSumDicValue, listIndex)

            # 列求和
            for key in colSumDicKey:
                isNeedColSum = titleItem.keyTrueInTitleItem(key)
                if isNeedColSum:
                    colSumDicCallFunc(key, 'setOperatorIndex', listIndex)

            colSumDicValue = titleItem.getColSumDicValue()
            isColSumDicResultHere = colSumDicValue is not None
            if isColSumDicResultHere:
                colSumDicCallFunc(colSumDicValue, 'setResultIndex', listIndex)
            listIndex = listIndex+1

    if listIndex == 0:
        maxlistIndex = listIndex
    else:
        maxlistIndex = listIndex-1

    # 日期&周目
    countNumCell_format = helper.getCountNumFormat()
    sum_format = helper.getSumResultFormat()
    monthEnd_format = helper.getMonthEndFormat()

    for c_month in range(max_month):
        c_month = c_month+1
        newWeek = True
        curdayIndex = 0
        montTotoalDayNum = calendar.monthrange(c_year, c_month)[1]
        weekStartLineIndex = weekEndLineIndex = 0
        montStartLineIndex = montEndLineIndex = 0
        if c_month == 1:
            lineIndex = lineIndex+jsonInfo.getTitleTotalLine()
        for dayItem in obj.itermonthdays4(c_year, c_month):
            if not dayItem[1] == c_month:
                continue
            if curdayIndex == 0:
                montStartLineIndex = lineIndex
            if newWeek:
                weekStartLineIndex = lineIndex
                newWeek = False
            # 日期：年月日，显示格式:x月x日
            yearMontDay = "%d-%d-%d" % (c_year, dayItem[1], dayItem[2])
            date = datetime.strptime(yearMontDay, "%Y-%m-%d")
            date_format = helper.getDateFormat()
            worksheet.write_datetime(lineIndex, 0, date, date_format)

            # 周数
            cnWeekNumStr = XlsHelper.getWeekDayCnNameByIndex(dayItem[3])
            weekday_format = helper.getWeekFormat()
            worksheet.write(lineIndex, 1, cnWeekNumStr, weekday_format)
            curdayIndex = curdayIndex+1

            # 行求和
            for sumItem in rowSumDic.keys():
                sumItemList = []
                for tmpLie in range(maxlistIndex):
                    for lie in countCellList:
                        if tmpLie == lie:
                            # 合计列内容
                            worksheet.write(lineIndex, tmpLie, "",
                                            countNumCell_format)
                    tmpList = rowSumDic[sumItem]
                    if isinstance(tmpList, RowSumResult):
                        for needCountLie in tmpList.itemList:
                            if tmpLie == needCountLie:
                                sumStr = XlsHelper.rowCol2Str(lineIndex, tmpLie)
                                sumItemList.append(sumStr)
                # 行求和内容
                sumStr = ",".join(sumItemList)
                sumStr = "=SUM(%s)" % sumStr
                result = rowSumDic[sumItem]
                if isinstance(result, RowSumResult):
                    worksheet.write(lineIndex, result.index, sumStr, sum_format)

            # 周求和
            weekCountKey = jsonInfo.getWeekCountKey()
            if cnWeekNumStr == '周日':
                weekEndLineIndex = lineIndex
                sumStr = getSumStr(weekCountKey, weekStartLineIndex, weekEndLineIndex)
                lineIndex = lineIndex+1
                result = colSumDic[weekCountKey]
                if isinstance(result, ColSumResult):
                    worksheet.write(lineIndex, result.resultIndex, sumStr, sum_format)
                newWeek = True
            #  月求和
            if curdayIndex == montTotoalDayNum:
                weekEndLineIndex = lineIndex
                montEndLineIndex = lineIndex

                if cnWeekNumStr != '周日':
                    sumStr = getSumStr(weekCountKey, weekStartLineIndex, weekEndLineIndex)
                    result = colSumDic[weekCountKey]
                    if isinstance(result, ColSumResult):
                        lineIndex = lineIndex+1
                        worksheet.write(lineIndex, result.resultIndex, sumStr, sum_format)

                monthCountKey = jsonInfo.getMonthCountKey()
                sumStr = getSumStr(monthCountKey, montStartLineIndex, montEndLineIndex)
                result = colSumDic[monthCountKey]
                if isinstance(result, ColSumResult):
                    lineIndex = lineIndex+1
                    for tmpCol in range(maxlistIndex):
                        if tmpCol < maxlistIndex:
                            worksheet.write(lineIndex, tmpCol, "", monthEnd_format)
                    worksheet.write(lineIndex, result.resultIndex, sumStr, monthEnd_format)
                newWeek = True
            # 行高
            defaultTitleHeight = jsonInfo.getTitleHeight()
            worksheet.set_row(lineIndex, defaultTitleHeight)
            lineIndex = lineIndex+1

    helper.close_workbook()


if __name__ == "__main__":
    main()
