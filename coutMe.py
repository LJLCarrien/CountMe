from datetime import datetime
import calendar
import json
import xlsxwriter
# 英文简称获取：calendar.day_abbr
cn_day_abbr = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
excelSavePath = './test.xlsx'
jsonFilePath = './countMe_data.json'
obj = calendar.Calendar()

c_year = datetime.now().year
c_month = 9

sumDic = {}
# 合计列下标数组
countCellList = []

workbook = xlsxwriter.Workbook(excelSavePath)
worksheet = workbook.add_worksheet(str(c_year))

date_format = workbook.add_format({'num_format': 'm"月"d"日"'})
date_format.set_align('center')
date_format.set_align('vcenter')

weekday_format = workbook.add_format()
weekday_format.set_align('left')
weekday_format.set_align('vcenter')

sum_format = workbook.add_format()
sum_format.set_bold()
sum_format.set_align('center')
sum_format.set_align('vcenter')
sum_format.set_font_size(16)
sum_format.set_font_name('宋体')
sum_format.set_font_color('red')

countNumCell_format = workbook.add_format()
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


class SumResult(object):
    def __init__(self):
        # 求和结果列下标
        self.index = -1
        # 用于求和的列下标
        self.itemList = []

    def setIndex(self, i):
        self.index = i

    def setList(self, list):
        self.itemList = list

    def addListItem(self, item):
        if self.itemList == None:
            self.itemList = []
        self.itemList.append(item)


def sumDicSetIndex(key, content):
    if key not in sumDic:
        sumDic[key] = SumResult()
    sumDic[key].setIndex(content)


def sumDicAddItem(key, listItem):
    if key not in sumDic:
        sumDic[key] = SumResult()
    sumDic[key].addListItem(listItem)


def main():
    with open(jsonFilePath, 'rb')as f:
        lineIndex = 0
        listIndex = 0

        jsonData = json.load(f)
        defaultTitleFontSize = jsonData['defaultTitleFontSize']
        defaultTitleFontName = jsonData['defaultTitleFontName']
        defaultSecTitleFontSize = jsonData['defaultSecTitleFontSize']
        secondTitle = jsonData['secondTitle']
        titleInfo = jsonData['title']
        sumDicKey = jsonData['sumDicKey']

        for titleItem in titleInfo:
            value = titleItem['name']
            bHaveSecMenu = 'secondSort' in titleItem

            tmp_title_format = getNewTitleFormat(
                titleItem, defaultTitleFontSize, defaultTitleFontName)

            # 含二级菜单
            if bHaveSecMenu:
                tmp_Sec_title_format = getNewTitleFormat(
                    titleItem, defaultSecTitleFontSize, defaultTitleFontName)

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
            rowHeight = jsonData['defaultTitleHeight']
            worksheet.set_row(lineIndex, rowHeight)

            if 'countCell' in titleItem:
                listIndex = listIndex+1
                countCellList.append(listIndex)
                worksheet.merge_range(lineIndex, listIndex, lineIndex+1, listIndex,
                                      titleItem['countCell'], tmp_title_format)

            if 'sumDic' in titleItem:
                key = titleItem['sumDic']
                sumDicSetIndex(key, listIndex)

            for key in sumDicKey:
                if key in titleItem:
                    sumDicAddItem(key, listIndex)

            listIndex = listIndex+1

    maxlistIndex = listIndex-1
    # 日期&周目
    lineIndex = lineIndex+1
    for day in obj.itermonthdays4(c_year, c_month):
        if not day[1] == c_month:
            continue

        # 日期：年月日，显示格式:x月x日
        yearMontDay = "%d-%d-%d" % (c_year, day[1], day[2])
        date = datetime.strptime(yearMontDay, "%Y-%m-%d")
        worksheet.write_datetime(lineIndex, 0, date, date_format)

        # 周数
        cnWeekNumStr = cn_day_abbr[day[3]]
        worksheet.write(lineIndex, 1, cnWeekNumStr, weekday_format)

        # 求和
        for sumItem in sumDic.keys():
            sumItemList = []
            for tmpLie in range(maxlistIndex):
                for lie in countCellList:
                    if tmpLie == lie:
                        worksheet.write(lineIndex, tmpLie, "", countNumCell_format)
                for needCountLie in sumDic[sumItem].itemList:
                    if tmpLie == needCountLie:
                        sumStr = rowCol2Str(lineIndex, tmpLie)
                        sumItemList.append(sumStr)
            sumStr = ",".join(sumItemList)
            sumStr = "=SUM(%s)" % sumStr
            worksheet.write(
                lineIndex, sumDic[sumItem].index, sumStr, sum_format)

        # 换行
        if cnWeekNumStr == '周日':
            lineIndex = lineIndex+1
        lineIndex = lineIndex+1

    workbook.close()

if __name__ == "__main__":
    main()
