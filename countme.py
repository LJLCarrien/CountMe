"""统计人生"""
from datetime import datetime
import calendar
import json
from configure_data import ConfigureData
from item_savedata import SaveDataItem
from sum_colsum import ColSumResult
from sum_rowsum import RowSumResult
from item_title import TitleItem
from helper_xlshelper import XlsHelper, get_jsondata
import operator

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
jsoninfo = get_jsondata(config_json_filepath)


def rowsum_set_resultindex(key, content):
  global rowsum_dic
  """
    行求和结果，列下标设置
    """
  if key not in rowsum_dic:
    rowsum_dic[key] = RowSumResult()
  tmp = rowsum_dic[key]
  if isinstance(tmp, RowSumResult):
    tmp.set_index(content)


def rowsum_add_itemindex(key, value):
  global rowsum_dic
  """
    行求和元素，列下标设置
    """
  if key not in rowsum_dic:
    rowsum_dic[key] = RowSumResult()
  tmp = rowsum_dic[key]
  if isinstance(tmp, RowSumResult):
    tmp.add_list_item(value)


def colsumdic_callfunc(key, funname, content):
  global colsum_dic
  """
    列求和字典,列下标设置
    """
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


def create_empty():
  global column_title_dic, title_list
  '''创建空excel模板'''
  excel_savepath = jsoninfo.get_excelpath()
  helper = XlsHelper(jsoninfo, excel_savepath)
  worksheet = helper.get_worksheet(str(c_year))

  title_list = jsoninfo.get_titlelist()

  rowsum_dickey = jsoninfo.get_rowsum_dickeys()
  colsum_dickey = jsoninfo.get_colsum_dickeys()

  lineindex = 0  # 行下标
  listindex = 0  # 列下标
  breeze_listindex = -1  # 冻结列下标

  for titleitem in title_list:
    if isinstance(titleitem, TitleItem):
      value_sditem = titleitem.showname

      titleformat = helper.get_titleformat(titleitem)

      secondmenu_list = titleitem.get_seclist()
      ishave_sectitle = secondmenu_list is not None
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
          column_title_dic[new_col] = titleitem.get_seclist_name_by_index(
              tmplie)
          worksheet.write(lineindex + 1, new_col,
                          titleitem.get_seclist_showname_by_index(tmplie),
                          sectitle_format)
      else:
        #标题列下标字典
        column_title_dic[listindex] = titleitem.name
        # 一级菜单设置列下标
        titleitem.set_columindex(listindex)
        worksheet.merge_range(lineindex, listindex, lineindex + 1, listindex,
                              value_sditem, titleformat)

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
      countcell_value = titleitem.get_countcell_value()
      is_countcell = countcell_value is not None
      if is_countcell:
        listindex = listindex + 1
        countcell_list.append(listindex)
        worksheet.merge_range(
            lineindex,
            listindex,
            lineindex + 1,
            listindex,
            countcell_value,
            titleformat,
        )

      # 行求和
      for key in rowsum_dickey:
        is_rowsum = titleitem.get_is_keytrue_in_titleitem(key)
        if is_rowsum:
          rowsum_add_itemindex(key, listindex)

      rowsumdic_value = titleitem.get_rowsum_dicvalue()
      is_rowsum_result_here = rowsumdic_value is not None
      if is_rowsum_result_here:
        rowsum_set_resultindex(rowsumdic_value, listindex)

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
      cn_weeknum_str = XlsHelper.get_weekday_cn_name_by_index(day_item[3])
      weekday_format = helper.get_weekformat()
      worksheet.write(lineindex, 1, cn_weeknum_str, weekday_format)
      curday_index = curday_index + 1

      #写入数据
      key_ymd = year_month_day.replace('-', '/')
      for key, value_sditem in savedata_dict.items():
        if key_ymd == key:
          for attr, attr_value in value_sditem.__dict__.items():
            titleitem = get_titleitem_by_name(attr)
            if titleitem is None:
              print(f'[error]这里出现意外的空标题，请检查列标题名{attr}')
              break
            if isinstance(attr_value, dict):
              # 合计值
              countcell_num = SaveDataItem.get_countcell_num(attr_value)
              # 有二级列表&有合计列 如：餐饮
              bhave_sec = titleitem.get_is_seclist()
              if bhave_sec:
                for v_key, v_value in attr_value.items():
                  for tmpcol in range(maxlist_index):
                    if tmpcol in column_title_dic and column_title_dic[
                        tmpcol] == v_key:
                      worksheet.write(lineindex, tmpcol, v_value)
                      break
                if countcell_num is not None:
                  coutcell_colindex = titleitem.get_countcell_col_index()
                  worksheet.write_number(lineindex, coutcell_colindex,
                                         countcell_num, count_numcell_format)
              else:
                # 一级列表
                des = SaveDataItem.get_des_contain_countcell(attr_value)
                for tmpcol in range(maxlist_index):
                  if tmpcol in column_title_dic and column_title_dic[
                      tmpcol] == attr:
                    # 一级列表&有合计列 如：生活用品、交通等
                    if des is not None:
                      worksheet.write(lineindex, tmpcol, des)
                    # 一级列表，内容直接就是合计 如：日必须，日合计
                    if countcell_num is not None:
                      coutcell_colindex = titleitem.get_countcell_col_index()
                      worksheet.write_number(lineindex, coutcell_colindex,
                                             countcell_num,
                                             count_numcell_format)
                    break
            else:
              # 无额外列，净内容 如：todo,havedone
              for tmpcol in range(maxlist_index):
                if tmpcol in column_title_dic and column_title_dic[
                    tmpcol] == attr:
                  worksheet.write(lineindex, tmpcol, attr_value)
                  break

      # 行求和
      for key, value_sditem in rowsum_dic.items():
        sum_itemlist = []
        for tmplie in range(maxlist_index):
          # for lie in countcell_list:
          #   if tmplie == lie:
          #     # 合计列内容,空表格的时候写入格式用的
          #     worksheet.write(lineindex, tmplie, "", count_numcell_format)
          if isinstance(value_sditem, RowSumResult):
            for needcount_lie in value_sditem.item_list:
              if tmplie == needcount_lie:
                sumstr = XlsHelper.get_rowcol_2_str(lineindex, tmplie)
                sum_itemlist.append(sumstr)
        # 行求和内容
        sumstr = ",".join(sum_itemlist)
        sumstr = f"=SUM({sumstr})"
        if isinstance(value_sditem, RowSumResult):
          worksheet.write(lineindex, value_sditem.index, sumstr, sum_format)

      # 周求和
      weekcount_key = jsoninfo.get_weekcount_key()
      if cn_weeknum_str == "周日":
        week_end_lineindex = lineindex
        sumstr = get_sumstr(weekcount_key, week_start_lineindex,
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

      # 行高
      defaulttitle_height = jsoninfo.get_title_height()
      worksheet.set_row(lineindex, defaulttitle_height)
      lineindex = lineindex + 1

  helper.close_workbook()


def read_savedata():
  with open(savedata_json_filepath, 'rb') as f:
    jsondic = json.load(f)
    for key, value in jsondic.items():
      if key == 'emptyTemplate':
        empty_template = value
      else:
        sd_item = SaveDataItem(empty_template)
        sd_item.set_data(value)
        savedata_dict[key] = sd_item


if __name__ == "__main__":
  read_savedata()
  create_empty()
