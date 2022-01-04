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
'''行求和字典'''
colsum_dic = {}
'''列求和字典'''

countcell_list = []
'''合计列下标数组'''

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


def colsumdic_callfunc(dic, key, funname, content):
  '''列求和字典,列下标设置'''
  if key not in dic:
    dic[key] = ColSumResult()
  mc = operator.methodcaller(funname, content)
  mc(dic[key])


def get_sumstr(dic, key, startindex, endindex):
  if key not in dic:
    dic[key] = ColSumResult()
  tmp = dic[key]
  if isinstance(tmp, ColSumResult):
    return tmp.get_sumstr_by_index(XlsHelper, startindex, endindex)
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
          colsumdic_callfunc(colsum_dic, key, "set_operatorindex", listindex)

      colsumdic_value = titleitem.get_colsum_dicvalue()
      is_colsum_dicresult_here = colsumdic_value is not None
      if is_colsum_dicresult_here:
        colsumdic_callfunc(colsum_dic, colsumdic_value, "set_resultindex",
                           listindex)
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
        if isinstance(value_sditem, RowSumResult):
          sumstr = value_sditem.get_sumstr(XlsHelper, lineindex)
          worksheet.write(lineindex, value_sditem.index, sumstr, sum_format)
          dataItem.cal_countcell(key, value_sditem.item_name_list)

      # 周求和
      weekcount_key = jsoninfo.get_weekcount_key()
      if cn_weeknum_str == "周日":
        week_end_lineindex = lineindex
        sumstr = get_sumstr(colsum_dic, weekcount_key, week_start_lineindex,
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
          sumstr = get_sumstr(colsum_dic, weekcount_key, week_start_lineindex,
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
        sumstr = get_sumstr(colsum_dic, monthcount_key, month_start_lineindex,
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


def read_savedata(filePath):
  global savedata_dict
  with open(filePath, 'rb') as f:
    jsondic = json.load(f)
    savedata_dict = jsondic


def write_analysis(jsoninfo: ConfigureData, helper: XlsHelper):
  '''分析结果'''
  sheetname = f"{str(c_year)}汇总"
  worksheet = helper.get_worksheet(sheetname)

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
    # 表格标题数据
    charttitle_row_start_index = row_cur + 1
    charttitle_row_end_index = row_cur + row_data_num

    charttitle_row_start_cell_abs = helper.get_rowcol_abs(
        charttitle_row_start_index, col_cur)
    charttitle_row_end_cell_abs = helper.get_rowcol_abs(
        charttitle_row_end_index, col_cur)

    chart_pos_cellstr = helper.get_rowcol_2_str(row_cur, col_cur + col_len + 1)

    chartname = f"{m+1}月"
    worksheet.write(row_cur, col_cur, chartname)
    chart = helper.get_chart('column')
    chart.set_title({'name': chartname})
    chart.set_size({'height': 200, 'width': 384})

    fontformat = {'color': '#ababb9', 'name': '微软雅黑', 'size': 9}
    lineformat = {'none': True}
    major_gridlines_format = {'visible': True, 'line': {'color': '#d9d9d9'}}
    chart.set_legend({'position': 'bottom', 'font': fontformat})
    chart.set_x_axis({'num_font': fontformat, 'line': lineformat})
    chart.set_y_axis({
        'num_font': fontformat,
        'major_gridlines': major_gridlines_format,
        'line': lineformat
    })
    #region 列标题
    for col_index in range(col_len):
      item = collist[col_index]
      charttitle_col_start_cellstr = helper.get_rowcol_2_str(
          charttitle_row_start_index - 1, col_cur + 1 + col_index)
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
              if row_index == 0:
                chartdata_col_start_cell_abs = helper.get_rowcol_abs(
                    row_cur + row_index + 1, col_need)
                chartdata_col_end_cell_abs = helper.get_rowcol_abs(
                    row_cur + row_index + 1 + col_data_num, col_need)
                chart_categories = f'={sheetname}!{charttitle_row_start_cell_abs}:{charttitle_row_end_cell_abs}'
                chart_name = f'={sheetname}!{charttitle_col_start_cellstr}'
                chart_value = f'={sheetname}!{chartdata_col_start_cell_abs}:{chartdata_col_end_cell_abs}'
                chart.add_series({
                    'categories': chart_categories,
                    'name': chart_name,
                    'values': chart_value,
                    'fill': {
                        'color': item.bgcolor
                    }
                })
          # 月求和
          if row_index == row_len - 1:
            rs = RowSumResult()
            for i in range(col_need - 1):
              if i == 0:
                continue
              rs.add_list_item(i)
            sumstr = rs.get_sumstr(XlsHelper, row_cur + row_index + 1)
            worksheet.write(row_cur + row_index + 1, col_need, sumstr,
                            itemcellformat)
      #endregion
    worksheet.insert_chart(chart_pos_cellstr, chart)
    if row_cur == 0:
      # 每月列数=列标题数+行标题占的1列
      col_content = col_need + 1

    #region 行标题
    for row_index in range(row_len):
      item = rowlist[row_index]
      if isinstance(item, RowAnalysisTitleItem):
        titleformat = helper.get_analysis_titleitem_format(item)
        # 避开第0行
        row_need = row_cur + row_index + 1
        # 行标题对应内容
        worksheet.write(row_need, col_cur, item.showname, titleformat)
        # 周求和
        isrow_weekcount = item.get_is_weekcount()
        if isrow_weekcount:
          is_handle_itemcell = item.get_isneed_handle_itemcell()
          if is_handle_itemcell:
            itemcellformat = helper.get_analysis_itemcell_format(item)
            for c in range(col_len):
              if c < col_content - 2:
                if c == col_content - 3:
                  worksheet.write(row_need, c + 1, None, itemcellformat)
                else:
                  cs = ColSumResult()
                  cs.set_index(row_need - 1 - 7, row_need - 2, c + 1, c + 1)
                  sumstr = cs.get_sumstr(XlsHelper)
                  worksheet.write(row_need, c + 1, sumstr, itemcellformat)
    #endregion

    if row_cur == 0:
      # 每月行数=行标题数+列标题占的1行
      row_content = row_need + 1

    row_cur = row_cur + row_content + row_interval

    col_index = 0


def write_chart_item(helper: XlsHelper, month):
  sheetname = f"{str(c_year)}图表"
  worksheet = helper.get_worksheet(sheetname)
  chart = helper.get_chart('column')
  chart.set_size({'height': 200, 'width': 384})
  chart_rowtitle = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
  chart_coltitle = ['第一周', '第二周', '第三周', '第四周', '第五周', '第六周']

  data = [
      [1, 2, 3, 4, 5, 6, 7],
      [2, 4, 6, 8, 10, 12, 14],
      [3, 6, 9, 12, 15, 18, 21],
      [4, 8, 12, 16, 20, 24, 28],
      [5, 10, 15, 20, 25, 30, 35],
      [6, 12, 18, 24, 30, 36, 42],
  ]

  worksheet.write_row('B1', chart_coltitle)
  worksheet.write_column('A2', chart_rowtitle)

  worksheet.write_column('B2', data[0])
  worksheet.write_column('C2', data[1])
  worksheet.write_column('D2', data[2])
  worksheet.write_column('E2', data[3])
  worksheet.write_column('F2', data[4])
  worksheet.write_column('G2', data[5])
  chart.add_series({
      'categories': f'={sheetname}!$A$2:$A$8',
      'name': f'={sheetname}!B1',
      'values': f'={sheetname}!$B$2:$B$8',
      'fill': {
          'color': '#5B9BD5'
      }
  })
  chart.add_series({
      'categories': f'={sheetname}!$A$2:$A$8',
      'name': f'={sheetname}!C1',
      'values': f'={sheetname}!$C$2:$C$8',
      'fill': {
          'color': '#ED7D31'
      }
  })
  chart.add_series({
      'categories': f'={sheetname}!$A$2:$A$8',
      'name': f'={sheetname}!D1',
      'values': f'={sheetname}!$D$2:$D$8',
      'fill': {
          'color': '#A5A5A5'
      }
  })
  chart.add_series({
      'categories': f'={sheetname}!$A$2:$A$8',
      'name': f'={sheetname}!E1',
      'values': f'={sheetname}!$E$2:$E$8',
      'fill': {
          'color': '#FFC000'
      }
  })
  chart.add_series({
      'categories': f'={sheetname}!$A$2:$A$8',
      'name': f'={sheetname}!F1',
      'values': f'={sheetname}!$F$2:$F$8',
      'fill': {
          'color': '#4472C0'
      }
  })
  chart.add_series({
      'categories': f'={sheetname}!$A$2:$A$8',
      'name': f'={sheetname}!G1',
      'values': f'={sheetname}!$G$2:$G$8',
      'fill': {
          'color': '#70AD47'
      }
  })
  chart.set_x_axis({
      'num_font': {
          'color': '#ababb9',
          'name': '微软雅黑',
          'size': 9
      },
      'line': {
          'none': True
      }
  })
  chart.set_y_axis({
      'num_font': {
          'color': '#ababb9',
          'name': '微软雅黑',
          'size': 9
      },
      'major_gridlines': {
          'visible': True,
          'line': {
              'color': '#d9d9d9'
          }
      },
      'line': {
          'none': True
      }
  })
  chart.set_title({'name': f'{month}月'})
  chart.set_legend({
      'position': 'bottom',
      'font': {
          'color': '#ababb9',
          'name': '微软雅黑',
          'size': 9
      }
  })
  worksheet.insert_chart('A9', chart)


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
  savedata_json_filepath = jsoninfo.get_savedata_json_filepath()

  helper = XlsHelper(jsoninfo, excel_savepath)

  read_savedata(savedata_json_filepath)
  create_detail(jsoninfo, helper)

  init_analysis(jsoninfo, month_lis_tuple_daydataItem_dic, row_daydataItem_dic)
  write_analysis(jsoninfo, helper)

  # 测试代码
  # write_chart_item(helper, 1)
  helper.close_workbook()
