from datetime import datetime
import calendar
import json
from xlsHelper import xlsHelper
import xlsxwriter
import operator
import jsonHelper
# 英文简称获取：calendar.day_abbr
cn_day_abbr = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
excelSavePath = './test.xlsx'
jsonFilePath = './countMe_data.json'
obj = calendar.Calendar()

c_year = datetime.now().year
c_month = 9

rowSumDic = {}
colSumDic = {}
# 合计列下标数组
countCellList = []

workbook = xlsxwriter.Workbook(excelSavePath)
worksheet = workbook.add_worksheet(str(c_year))

date_format = workbook.add_format({'num_format': 'm"月"d"日"'})
date_format.set_font_name('宋体')
date_format.set_align('center')
date_format.set_align('vcenter')

weekday_format = workbook.add_format()
date_format.set_font_name('宋体')
weekday_format.set_align('left')
weekday_format.set_align('vcenter')

sum_format = workbook.add_format()
sum_format.set_bold()
sum_format.set_align('center')
sum_format.set_align('vcenter')
sum_format.set_font_size(16)
sum_format.set_font_name('宋体')
sum_format.set_font_color('red')

# 合计
countNumCell_format = workbook.add_format({'num_format': '0.00_'})
countNumCell_format.set_align('center')
countNumCell_format.set_align('vcenter')
countNumCell_format.set_font_size(11)
countNumCell_format.set_font_name('宋体')
countNumCell_format.set_font_color('red')

# print(ord('A'))
# print(chr(65))


def rowCol2Str(row, col):
    # 行列转字符串
    # eq: 0,0->A1 ; 1,0->A2 ; 0，1->B1 ;
    result = "%s%s" % (chr(ord('A')+col), row+1)
    return result


def getNewTitleFormat(itemInfo, fontSize, fontName):
    format = workbook.add_format()
    format.set_bold()
    format.set_align('center')
    format.set_align('vcenter')
    format.set_font_size(fontSize)
    format.set_font_name(fontName)
    if 'fontColor' in itemInfo:
        format.set_font_color(itemInfo['fontColor'])
    if 'bgColor' in itemInfo:
        format.set_bg_color(itemInfo['bgColor'])
    return format


class RowSumResult(object):
    def __init__(self):
        # 求和结果-列下标
        self.index = -1
        # 用于求和-列下标
        self.itemList = []

    def setIndex(self, i):
        self.index = i

    def setList(self, list):
        self.itemList = list

    def addListItem(self, item):
        if self.itemList == None:
            self.itemList = []
        self.itemList.append(item)


def rowSumDicSetIndex(key, content):
    '''
    行求和字典,列下标设置
    '''
    if key not in rowSumDic:
        rowSumDic[key] = RowSumResult()
    rowSumDic[key].setIndex(content)


def rowSumDicAddItem(key, listItem):
    if key not in rowSumDic:
        rowSumDic[key] = RowSumResult()
    rowSumDic[key].addListItem(listItem)


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
            beginStr = rowCol2Str(self.rowStartIndex, self.OperatorIndex)
            endStr = rowCol2Str(self.rowEndIndex, self.OperatorIndex)
            return "=SUM(%s:%s)" % (beginStr, endStr)
        return ""


def colSumDicSetIndex(key, funName, content):
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
    return colSumDic[key].getSumStr(startIndex, endIndex)


def main():
    with open(jsonFilePath, 'rb')as f:
        lineIndex = 0
        listIndex = 0

        jsonData = json.load(f)
        titleFontSize = jsonData['titleFontSize']
        titleFontName = jsonData['titleFontName']
        secTitleFontSize = jsonData['secTitleFontSize']
        secondTitle = jsonData['secondTitle']
        titleInfo = jsonData['title']
        defaultTitleHeight = jsonData['defaultTitleHeight']
        # 行求和
        rowSumDicKey = jsonData['rowSumDicKey']
        colSumDicKey = jsonData['colSumDicKey']

        for titleItem in titleInfo:
            value = titleItem['name']
            bHaveSecMenu = 'secondSort' in titleItem

            tmp_title_format = getNewTitleFormat(
                titleItem, titleFontSize, titleFontName)

            # 含二级菜单
            if bHaveSecMenu:
                tmp_Sec_title_format = getNewTitleFormat(
                    titleItem, secTitleFontSize, titleFontName)

                oldlistIndex = listIndex
                secondMenuList = secondTitle[titleItem['secondSort']]
                secondMenuLen = len(secondMenuList)
                listIndex = listIndex+secondMenuLen-1
                # 一级菜单合并-行列行列
                worksheet.merge_range(
                    lineIndex, oldlistIndex, lineIndex, listIndex, value, tmp_title_format)
                # 二级菜单内容
                for tmpLie in range(secondMenuLen):
                    worksheet.write(lineIndex+1, oldlistIndex + tmpLie,
                                    secondMenuList[tmpLie], tmp_Sec_title_format)
            else:
                # worksheet.write(lineIndex, listIndex, value, tmp_title_format)
                worksheet.merge_range(
                    lineIndex, listIndex, lineIndex+1, listIndex, value, tmp_title_format)

            # 列宽
            colWidth = jsonData['defaultTitleWidth']
            if 'width' in titleItem:
                colWidth = titleItem['width']
            # 开始列，结束列，列宽
            worksheet.set_column(listIndex, listIndex, colWidth)

            # 行高
            worksheet.set_row(lineIndex, defaultTitleHeight)

            if 'countCell' in titleItem:
                listIndex = listIndex+1
                countCellList.append(listIndex)
                worksheet.merge_range(lineIndex, listIndex, lineIndex+1, listIndex,
                                      titleItem['countCell'], tmp_title_format)
            '''
            行求和
            '''
            if 'rowSumDic' in titleItem:
                key = titleItem['rowSumDic']
                rowSumDicSetIndex(key, listIndex)

            for key in rowSumDicKey:
                if key in titleItem and titleItem[key]:
                    rowSumDicAddItem(key, listIndex)

            '''
            列求和
            '''
            if 'colSumDic' in titleItem:
                key = titleItem['colSumDic']
                colSumDicSetIndex(key, 'setResultIndex', listIndex)

            for key in colSumDicKey:
                if key in titleItem and titleItem[key]:
                    colSumDicSetIndex(key, 'setOperatorIndex', listIndex)
            listIndex = listIndex+1
    maxlistIndex = listIndex-1

    # 日期&周目
    newWeek = True
    curdayIndex = 0
    montTotoalDayNum = calendar.monthrange(c_year, c_month)[1]

    weekStartLineIndex = 0
    weekEndLineIndex = 0

    montStartLineIndex = 0
    montEndLineIndex = 0

    lineIndex = lineIndex+1+len(secondTitle)
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
        worksheet.write_datetime(lineIndex, 0, date, date_format)

        # 周数
        cnWeekNumStr = cn_day_abbr[dayItem[3]]
        worksheet.write(lineIndex, 1, cnWeekNumStr, weekday_format)
        curdayIndex = curdayIndex+1

        # 行求和
        for sumItem in rowSumDic.keys():
            sumItemList = []
            for tmpLie in range(maxlistIndex):
                for lie in countCellList:
                    if tmpLie == lie:
                        # 求和格式
                        worksheet.write(lineIndex, tmpLie, "",countNumCell_format)
                for needCountLie in rowSumDic[sumItem].itemList:
                    if tmpLie == needCountLie:
                        sumStr = rowCol2Str(lineIndex, tmpLie)
                        sumItemList.append(sumStr)
            sumStr = ",".join(sumItemList)
            sumStr = "=SUM(%s)" % sumStr
            worksheet.write(
                lineIndex, rowSumDic[sumItem].index, sumStr, sum_format)

        # 周求和
        if cnWeekNumStr == '周日':
            weekEndLineIndex = lineIndex
            sumStr = getSumStr('weekCount', weekStartLineIndex, weekEndLineIndex)
            lineIndex = lineIndex+1
            worksheet.write(lineIndex, colSumDic['weekCount'].resultIndex, sumStr, sum_format)
            newWeek = True
        #  月求和
        if curdayIndex == montTotoalDayNum:
            weekEndLineIndex = lineIndex
            montEndLineIndex = lineIndex

            if cnWeekNumStr != '周日':
                sumStr = getSumStr('weekCount', weekStartLineIndex, weekEndLineIndex)
                lineIndex = lineIndex+1
                worksheet.write(lineIndex, colSumDic['weekCount'].resultIndex, sumStr, sum_format)

            sumStr = getSumStr('weekCount', montStartLineIndex, montEndLineIndex)
            lineIndex = lineIndex+1
            worksheet.write(
                lineIndex, colSumDic['MonthCount'].resultIndex, sumStr, sum_format)
            newWeek = True
        worksheet.set_row(lineIndex, defaultTitleHeight)
        lineIndex = lineIndex+1

    workbook.close()


def newMain():
    jsonInfo = jsonHelper.getJsonData(jsonFilePath)
    helper = xlsHelper(jsonInfo, excelSavePath)
    worksheet = helper.get_worksheet(str(c_year))
    titleList = jsonInfo.getTitleList()

    lineIndex = 0  # 行下标
    listIndex = 0  # 列下标
    for titleItem in titleList:
        if isinstance(titleItem, jsonHelper.TitleItem):
            value = titleItem.name
            titleFormat = helper.getNewTitleFormat(titleItem)

            secondMenuList = jsonInfo.getSecTitle(titleItem)
            isHaveSecTitle = secondMenuList is not None
            if isHaveSecTitle:
                oldlistIndex = listIndex
                secondMenuLen = len(secondMenuList)
                listIndex = listIndex+secondMenuLen-1
                # 一级菜单合并-行列行列
                worksheet.merge_range(
                    lineIndex, oldlistIndex, lineIndex, listIndex, value, titleFormat)
                # 二级菜单内容
                SecTitleFormat = helper.getNewTitleFormat(titleItem, True)
                for tmpLie in range(secondMenuLen):
                    worksheet.write(lineIndex+1, oldlistIndex + tmpLie,
                                    secondMenuList[tmpLie], SecTitleFormat)
            else:
                worksheet.merge_range(
                    lineIndex, listIndex, lineIndex+1, listIndex, value, titleFormat)
            # 列宽行高
            titleWidth = jsonInfo.getTitleWidth(titleItem)
            worksheet.set_column(listIndex, listIndex, titleWidth)
            titleHeight = jsonInfo.getTitleHeight()
            worksheet.set_row(lineIndex, titleHeight)

            CountCellValue = jsonInfo.getCountCell(titleItem)
            isHaveCountCell = CountCellValue is not None
            if isHaveCountCell:
                listIndex = listIndex+1
                countCellList.append(listIndex)
                worksheet.merge_range(lineIndex, listIndex, lineIndex+1, listIndex,
                                      CountCellValue, titleFormat)
    helper.close_workbook()

if __name__ == "__main__":
    # main()
    newMain()
