class MonthDataItem():
  def __init__(self, month) -> None:
    self.month = month
    self.daytuple_list = []
    '''天号数：[(1,3),(4,10),(11,17),(18,24),(25,31)]'''
    self.rowtuple_list = []
    '''行号数：[(1,3),(4,10),(11,17),(18,24),(25,31)]'''

  def add_weekitem_day(self, week_start_day, week_end_day):
    '''week_start_day：周开始日；week_end_day：结束日'''
    result = (week_start_day, week_end_day)
    self.daytuple_list.append(result)

  def add_weekitem_row(self, week_start_row, week_end_row):
    '''week_start_row：周开始行；week_end_row：结束行'''
    result = (week_start_row, week_end_row)
    self.rowtuple_list.append(result)


class AnalysisMonthDataItem():
  def __init__(
      self,
      index: int,
      table_rownum,
      table_colnum,
  ) -> None:
    self.data = {}
    self.monthdata: MonthDataItem = None
    self.index = index
    self.month = index + 1
    self.showtable_title = f"{self.month}月"
    '''月份标题'''
    self.showtable_weeknum = table_colnum
    '''用于表格显示的：周数'''
    self.showtable_weekdaynum = table_rownum
    '''用于表格显示的：周目数'''
    self.weekcoust_list = []
    '''周合计列表'''
    self.monthcoust = 0
    '''月合计'''
    self.rest_data()

  def rest_data(self):
    '''      j:[0],[1]......[6]
    [0]第一周：周一，周二....周日
    至
    [5]第六周：周一，周二....周日
    '''
    for i in range(self.showtable_weeknum):
      weeklist = []
      for j in range(self.showtable_weekdaynum):
        weeklist.append(None)
      self.data[i] = weeklist

  def set_monthdata(self, data: MonthDataItem):
    self.monthdata = data

  def set_data(self, weeknum, weekdaynum, value):
    self.data[weeknum][weekdaynum] = value

  def add_weekcoust_list(self, value):
    self.weekcoust_list.append(value)
    self.monthcoust = self.monthcoust + value

  # def set_data(self, data:MonthDataItem):
  #   weekindex = 0
  #   for daytuple in daytuple_list:
  #     week_daylist = list(range(daytuple))
  #     self.data[weekindex]=
  #     weekindex=weekindex+1
