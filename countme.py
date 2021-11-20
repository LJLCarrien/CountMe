"""统计人生"""
from datetime import datetime
import calendar
import json
from configure_data import ConfigureData
from item_analysis_month_data import AnalysisMonthDataItem, MonthDataItem
from item_analysis_title import AnalysisTitleType, ColAnalysisTitleItem, RowAnalysisTitleItem
from item_data import DataItem
from sum_colsum import ColSumResult
from sum_rowsum import RowSumResult
from item_title import TitleItem
from helper_xlshelper import XlsHelper, get_jsondata
import operator
import copy

obj = calendar.Calendar()

c_year = datetime.now().year
max_month = 12

rowsum_dic = {}
colsum_dic = {}
# 合计列下标数组
countcell_list = []

savedata_json_filepath = "./save_data_for_test.json"
savedata_dict = {}
column_title_dic = {}

title_list = []
config_json_filepath = "./config.json"

analy_monthdata_list = []

row_daydataItem_dic = {}
'''行下标-每天数据-字典；
key：行下标
value：DataItem'''

month_lis_tuple_daydataItem_dic = {}
'''月份数-每天数据列表-字典；
key：月份数；
value：MonthDataItem
'''


def get_day_by_row(row):
  '''通过行号获取日期day'''
  global row_daydataItem_dic
  if row in row_daydataItem_dic:
    item = row_daydataItem_dic[row]
    if isinstance(item, DataItem):
      return item.day
  return None


def rowsum_set_result(name, index):
  '''行求和结果，列下标设置'''
  global rowsum_dic
  if name not in rowsum_dic:
    rowsum_dic[name] = RowSumResult()
  tmp = rowsum_dic[name]
  if isinstance(tmp, RowSumResult):
    tmp.set_index(index)
    tmp.set_name(name)


def rowsum_add_item(key, index, name):
  global rowsum_dic
  """
    行求和元素，列下标设置
    """
  if key not in rowsum_dic:
    rowsum_dic[key] = RowSumResult()
  tmp = rowsum_dic[key]
  if isinstance(tmp, RowSumResult):
    tmp.add_list_item(index)
    tmp.add_list_item_name(name)


def colsumdic_callfunc(key, funname, content):
  '''列求和字典,列下标设置'''
  global colsum_dic
  if key not in colsum_dic:
    colsum_dic[key] = ColSumResult()
  mc = operator.methodcaller(funname, content)
  mc(colsum_dic[key])


def get_sumstr(key, startindex, endindex):
  global colsum_dic
  if key not in colsum_dic:
    colsum_dic[key] = ColSumResult()
  tmp = colsum_dic[key]
  if isinstance(tmp, ColSumResult):
    return tmp.get_sumstr(XlsHelper, startindex, endindex)
  return None


def get_titleitem_by_name(titlename: str) -> TitleItem:
  global title_list
  if title_list is None or len(title_list) == 0:
    return None
  for item in title_list:
    if isinstance(item, TitleItem) and item.name == titlename:
      return item
  return None


def create_detail(jsoninfo: ConfigureData, helper: XlsHelper):
  global column_title_dic, title_list
  worksheet = helper.get_worksheet(str(c_year))

  title_list = jsoninfo.get_titlelist()

  rowsum_dickey = jsoninfo.get_rowsum_dickeys()
  colsum_dickey = jsoninfo.get_colsum_dickeys()

  lineindex = 0  # 行下标
  listindex = 0  # 列下标
  breeze_listindex = -1  # 冻结列下标

  empty_dataItem = DataItem()

  for titleitem in title_list:
    if isinstance(titleitem, TitleItem):
      value_sditem = titleitem.showname

      titleformat = helper.get_titleformat(titleitem)

      secondmenu_list = titleitem.get_seclist()
      ishave_sectitle = secondmenu_list is not None

      addcountcell_value = titleitem.get_addcountcell_value()
      is_addcountcell = addcountcell_value is not None
      is_countcell = titleitem.get_countcell()
      if ishave_sectitle:

        oldlist_index = listindex
        secondmenu_len = len(secondmenu_list)
        listindex = listindex + secondmenu_len - 1
        # 一级菜单设置列下标
        for i in range(oldlist_index, listindex + 1):
          titleitem.set_columindex(i)
        # 一级菜单合并-行列行列
        worksheet.merge_range(lineindex, oldlist_index, lineindex, listindex,
                              value_sditem, titleformat)
        # 二级菜单内容
        sectitle_format = helper.get_titleformat(titleitem, True)
        for tmplie in range(secondmenu_len):
          new_col = oldlist_index + tmplie
          #标题列下标字典
          sectitle_name = titleitem.get_seclist_name_by_index(tmplie)
          sectitle_showname = titleitem.get_seclist_showname_by_index(tmplie)
          column_title_dic[new_col] = sectitle_name
          worksheet.write(lineindex + 1, new_col, sectitle_showname,
                          sectitle_format)
          empty_dataItem.set_data_property(titleitem.name, sectitle_name)
      else:
        #标题列下标字典
        column_title_dic[listindex] = titleitem.name
        # 一级菜单设置列下标
        titleitem.set_columindex(listindex)
        worksheet.merge_range(lineindex, listindex, lineindex + 1, listindex,
                              value_sditem, titleformat)
        if is_addcountcell:
          empty_dataItem.set_data_property(titleitem.name)
        else:
          if is_countcell:
            empty_dataItem.set_data_property_countcell(titleitem.name)
          else:
            empty_dataItem.set_data(titleitem.name)

      # 列宽行高
      title_width = titleitem.get_title_width()
      worksheet.set_column(listindex, listindex, title_width)
      title_height = jsoninfo.get_title_height()
      worksheet.set_row(lineindex, title_height)

      # 冻结
      is_freeze = titleitem.get_isfreeze()
      if is_freeze:
        breeze_listindex = listindex
      # 合计

      if is_addcountcell:
        empty_dataItem.set_data_property_countcell(titleitem.name)

        listindex = listindex + 1
        countcell_list.append(listindex)
        worksheet.merge_range(
            lineindex,
            listindex,
            lineindex + 1,
            listindex,
            addcountcell_value,
            titleformat,
        )

      # 行求和
      for key in rowsum_dickey:
        is_rowsum = titleitem.get_is_keytrue_in_titleitem(key)
        if is_rowsum:
          rowsum_add_item(key, listindex, titleitem.name)

      rowsumdic_value = titleitem.get_rowsum_dicvalue()
      is_rowsum_result_here = rowsumdic_value is not None
      if is_rowsum_result_here:
        rowsum_set_result(rowsumdic_value, listindex)

      # 列求和
      for key in colsum_dickey:
        is_colsum = titleitem.get_is_keytrue_in_titleitem(key)
        if is_colsum:
          colsumdic_callfunc(key, "set_operatorindex", listindex)

      colsumdic_value = titleitem.get_colsum_dicvalue()
      is_colsum_dicresult_here = colsumdic_value is not None
      if is_colsum_dicresult_here:
        colsumdic_callfunc(colsumdic_value, "set_resultindex", listindex)
      listindex = listindex + 1

  # print(empty_dataItem)

  if listindex == 0:
    maxlist_index = listindex
  else:
    maxlist_index = listindex - 1

  # 日期&周目
  count_numcell_format = helper.get_countnum_format()
  sum_format = helper.get_sumresult_format()
  month_end_format = helper.get_monthend_format()

  lineindex = lineindex + jsoninfo.get_titletotal_line()
  first_showday_line = lineindex
  # 按月/行冻结
  freeze_detail = jsoninfo.get_freezeline_detail()
  # 默认冻结
  if breeze_listindex != -1:
    worksheet.freeze_panes(first_showday_line, breeze_listindex,
                           first_showday_line, 0)
  for c_month in range(max_month):
    c_month = c_month + 1
    new_week = True
    curday_index = 0
    month_totoal_daynum = calendar.monthrange(c_year, c_month)[1]
    week_start_lineindex = week_end_lineindex = 0
    month_start_lineindex = month_end_lineindex = 0

    month_dataitem = MonthDataItem(c_month)
    month_lis_tuple_daydataItem_dic[c_month] = month_dataitem

    # 按月冻结行
    if (breeze_listindex != -1 and int(freeze_detail) == c_month
        and jsoninfo.get_isfreezemonth()):
      worksheet.freeze_panes(first_showday_line, breeze_listindex, lineindex,
                             0)
    for day_item in obj.itermonthdays4(c_year, c_month):
      if not day_item[1] == c_month:
        continue
      if curday_index == 0:
        month_start_lineindex = lineindex
      if new_week:
        week_start_lineindex = lineindex
        new_week = False
      # 按指定行冻结行
      if (breeze_listindex != -1 and int(freeze_detail) - 1 == lineindex
          and jsoninfo.get_isfreezeline()):
        worksheet.freeze_panes(first_showday_line, breeze_listindex, lineindex,
                               0)
      # 日期：年月日，显示格式:x月x日
      year_month_day = f"{c_year}-{day_item[1]}-{day_item[2]}"
      date = datetime.strptime(year_month_day, "%Y-%m-%d")
      date_format = helper.get_dateformat()
      worksheet.write_datetime(lineindex, 0, date, date_format)

      # 周数
      weekday_index = day_item[3]
      cn_weeknum_str = XlsHelper.get_weekday_cn_name_by_index(weekday_index)
      weekday_format = helper.get_weekformat()
      worksheet.write(lineindex, 1, cn_weeknum_str, weekday_format)
      curday_index = curday_index + 1

      #写入数据
      key_ymd = year_month_day.replace('-', '/')
      dataItem = copy.deepcopy(empty_dataItem)
      dataItem.row = lineindex
      dataItem.date = key_ymd
      dataItem.set_weekday_index(weekday_index)
      row_daydataItem_dic[lineindex] = dataItem

      for key, value_sditem in savedata_dict.items():
        if key_ymd == key and isinstance(value_sditem, dict):
          for attr, attr_value in value_sditem.items():
            titleitem = get_titleitem_by_name(attr)
            name = titleitem.name
            is_countcell = titleitem.get_countcell()
            if titleitem is None:
              print(f'[error]这里出现意外的空标题，请检查列标题名{attr}')
              break
            if isinstance(attr_value, dict):
              # 合计值
              countcell_num = DataItem.get_countcell_num(attr_value)
              # 有二级列表&有合计列 如：餐饮
              bhave_sec = titleitem.get_is_seclist()
              if bhave_sec:
                for v_key, v_value in attr_value.items():
                  for tmpcol in range(maxlist_index):
                    if tmpcol in column_title_dic and column_title_dic[
                        tmpcol] == v_key:
                      worksheet.write(lineindex, tmpcol, v_value)
                      dataItem.set_data_property(name, v_key, v_value)
                      break
                if countcell_num is not None:
                  coutcell_colindex = titleitem.get_countcell_col_index()
                  worksheet.write_number(lineindex, coutcell_colindex,
                                         countcell_num, count_numcell_format)
                  dataItem.set_data_property_countcell(name, countcell_num)
              else:
                # 一级列表
                des = DataItem.get_des_contain_countcell(attr_value)
                for tmpcol in range(maxlist_index):
                  if tmpcol in column_title_dic and column_title_dic[
                      tmpcol] == attr:
                    if des is not None:
                      worksheet.write(lineindex, tmpcol, des)
                    # 一级列表&有合计列 如：生活用品、交通等
                    if countcell_num is not None:
                      coutcell_colindex = titleitem.get_countcell_col_index()
                      worksheet.write_number(lineindex, coutcell_colindex,
                                             countcell_num,
                                             count_numcell_format)
                      if not is_countcell:
                        dataItem.set_des(name, des)
                      dataItem.set_data_property_countcell(name, countcell_num)
                    # 一级列表，内容直接就是合计 如：日必须，日合计
                    else:
                      dataItem.set_data_property_countcell(name, countcell_num)

                    break
            else:
              # 无额外列，净内容 如：todo,havedone
              for tmpcol in range(maxlist_index):
                if tmpcol in column_title_dic and column_title_dic[
                    tmpcol] == attr:
                  worksheet.write(lineindex, tmpcol, attr_value)
                  dataItem.set_data(name, attr_value)
                  break
      # print(dataItem)
      # 行求和
      for key, value_sditem in rowsum_dic.items():
        sum_itemlist = []
        for tmplie in range(maxlist_index):
          # for lie in countcell_list:
          #   if tmplie == lie:
          #     # 合计列内容,空表格的时候写入格式用的
          #     worksheet.write(lineindex, tmplie, "", count_numcell_format)
          if isinstance(value_sditem, RowSumResult):
            for needcount_lie in value_sditem.item_index_list:
              if tmplie == needcount_lie:
                sumstr = XlsHelper.get_rowcol_2_str(lineindex, tmplie)
                sum_itemlist.append(sumstr)
        # 行求和内容
        sumstr = ",".join(sum_itemlist)
        sumstr = f"=SUM({sumstr})"
        if isinstance(value_sditem, RowSumResult):
          worksheet.write(lineindex, value_sditem.index, sumstr, sum_format)
          dataItem.cal_countcell(key, value_sditem.item_name_list)

      # 周求和
      weekcount_key = jsoninfo.get_weekcount_key()
      if cn_weeknum_str == "周日":
        week_end_lineindex = lineindex
        sumstr = get_sumstr(weekcount_key, week_start_lineindex,
                            week_end_lineindex)
        week_start_day = get_day_by_row(week_start_lineindex)
        week_end_day = get_day_by_row(week_end_lineindex)
        month_dataitem.add_weekitem_day(week_start_day, week_end_day)
        month_dataitem.add_weekitem_row(week_start_lineindex,
                                        week_end_lineindex)
        lineindex = lineindex + 1
        result = colsum_dic[weekcount_key]
        if isinstance(result, ColSumResult):
          worksheet.write(lineindex, result.result_index, sumstr, sum_format)
        new_week = True
      #  月求和
      if curday_index == month_totoal_daynum:
        week_end_lineindex = lineindex
        month_end_lineindex = lineindex

        if cn_weeknum_str != "周日":
          sumstr = get_sumstr(weekcount_key, week_start_lineindex,
                              week_end_lineindex)
          week_start_day = get_day_by_row(week_start_lineindex)
          week_end_day = get_day_by_row(week_end_lineindex)
          month_dataitem.add_weekitem_day(week_start_day, week_end_day)
          month_dataitem.add_weekitem_row(week_start_lineindex,
                                          week_end_lineindex)
          result = colsum_dic[weekcount_key]
          if isinstance(result, ColSumResult):
            lineindex = lineindex + 1
            worksheet.write(lineindex, result.result_index, sumstr, sum_format)

        monthcount_key = jsoninfo.get_monthcount_key()
        sumstr = get_sumstr(monthcount_key, month_start_lineindex,
                            month_end_lineindex)
        result = colsum_dic[monthcount_key]
        if isinstance(result, ColSumResult):
          lineindex = lineindex + 1
          for tmpcol in range(maxlist_index):
            if tmpcol < maxlist_index:
              worksheet.write(lineindex, tmpcol, "", month_end_format)
          worksheet.write(lineindex, result.result_index, sumstr,
                          month_end_format)
        new_week = True

      # print(row_daydataItem_dic)
      # 行高
      defaulttitle_height = jsoninfo.get_title_height()
      worksheet.set_row(lineindex, defaulttitle_height)
      lineindex = lineindex + 1


def read_savedata():
  global savedata_dict
  with open(savedata_json_filepath, 'rb') as f:
    jsondic = json.load(f)
    savedata_dict = jsondic


def write_analysis(jsoninfo: ConfigureData, helper: XlsHelper):
  '''分析结果'''
  worksheet = helper.get_worksheet(f"{str(c_year)}汇总")

  rowlist = jsoninfo.get_analysis_rowtitle_list()
  row_len = len(rowlist)
  collist = jsoninfo.get_analysis_coltitle_list()
  col_len = len(collist)
  row_cur = col_cur = 0
  row_data_num, col_data_num = jsoninfo.get_analysis_data_row_col_num()
  # 每个月内容行数
  row_content = 0
  col_content = 0
  # 每个月之间的间距
  row_interval = 1
  for m in range(max_month):
    worksheet.write(row_cur, col_cur, f"{m+1}月")
    # 列标题
    for col_index in range(col_len):
      item = collist[col_index]
      if isinstance(item, ColAnalysisTitleItem):
        titleformat = helper.get_analysis_titleitem_format(item)
        # 避开第0列
        col_need = col_index + 1
        worksheet.write(row_cur, col_need, item.showname, titleformat)
        # 列标题对应内容
        is_handle_itemcell = item.get_isneed_handle_itemcell()
        itemcellformat = helper.get_analysis_itemcell_format(
            item) if is_handle_itemcell else None
        for row_index in range(row_len):
          monthdata = analy_monthdata_list[m]
          if isinstance(monthdata, AnalysisMonthDataItem):
            if col_index < col_data_num and row_index < row_data_num:
              value = monthdata.get_data(col_index, row_index)
              worksheet.write(row_cur + row_index + 1, col_need, value,
                              itemcellformat)
          if row_index == row_len - 1:
            worksheet.write_number(row_cur + row_index + 1, col_need, 0,
                                   itemcellformat)
            # else:
            #   worksheet.write(row_cur + r + 1, col_need, None, itemcellformat)

    if row_cur == 0:
      # 每月列数=列标题数+行标题占的1列
      col_content = col_need + 1

    # 行标题
    for row_index in range(row_len):
      item = rowlist[row_index]
      if isinstance(item, RowAnalysisTitleItem):
        titleformat = helper.get_analysis_titleitem_format(item)
        # 避开第0行
        row_need = row_cur + row_index + 1
        worksheet.write(row_need, col_cur, item.showname, titleformat)
        # 行标题对应内容
        is_handle_itemcell = item.get_isneed_handle_itemcell()
        if is_handle_itemcell:
          itemcellformat = helper.get_analysis_itemcell_format(item)
          for c in range(col_len):
            if c < col_content - 2:
              if c == col_content - 3:
                worksheet.write(row_need, c + 1, None, itemcellformat)
              else:
                worksheet.write_number(row_need, c + 1, 0, itemcellformat)

    if row_cur == 0:
      # 每月行数=行标题数+列标题占的1行
      row_content = row_need + 1

    row_cur = row_cur + row_content + row_interval
    col_index = 0


def init_analysis(jsoninfo: ConfigureData, month_sort_dic, dataItem_dic):
  num = jsoninfo.get_analysis_data_row_col_num()
  row, col = num[0], num[1]
  for month_index in range(max_month):
    analy_monthdataitem = AnalysisMonthDataItem(month_index, row, col)
    month = analy_monthdataitem.month
    if month in month_sort_dic:
      monthdata_item = month_sort_dic[month]
      if isinstance(monthdata_item, MonthDataItem):
        analy_monthdataitem.set_monthdata(monthdata_item)
        rowtuple_list = monthdata_item.rowtuple_list
        for weekindex, rowtuple in enumerate(rowtuple_list):
          week_start_row = rowtuple[0]
          week_end_row = rowtuple[1]
          row_list = list(range(week_start_row, week_end_row + 1))
          weektotalcount = 0
          for dayrow in row_list:
            if dayrow in dataItem_dic:
              dataitem = dataItem_dic[dayrow]
              if isinstance(dataitem, DataItem):
                daycoust = dataitem.get_daycoust_countcell()
                weekdayindex = dataitem.get_weekday_index()
                analy_monthdataitem.set_data(weekindex, weekdayindex, daycoust)
                weektotalcount = weektotalcount + daycoust
          analy_monthdataitem.set_weekcoust_item(weekindex, weektotalcount)

    analy_monthdata_list.append(analy_monthdataitem)
  # print(analy_monthdata_list)


if __name__ == "__main__":
  jsoninfo = get_jsondata(config_json_filepath)
  excel_savepath = jsoninfo.get_excelpath()
  helper = XlsHelper(jsoninfo, excel_savepath)

  read_savedata()
  create_detail(jsoninfo, helper)

  init_analysis(jsoninfo, month_lis_tuple_daydataItem_dic, row_daydataItem_dic)
  write_analysis(jsoninfo, helper)

  helper.close_workbook()
