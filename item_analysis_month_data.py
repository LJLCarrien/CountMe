class AnalysisMonthDataItem():
  def __init__(self, index: int, rownum, colnum) -> None:
    self.data = {}
    self.index = index
    self.tableTitle = f"{self.index+1}月"
    '''月份'''
    self.weeknum = colnum
    '''周数'''
    self.weekdaynum = rownum
    '''数目数'''
    self.rest_data()

  def rest_data(self):
    '''      j:[0],[1]......[6]
    [0]第一周：周一，周二....周日
    至
    [5]第六周：周一，周二....周日
    '''
    for i in range(self.weeknum):
      weeklist = []
      for j in range(self.weekdaynum):
        weeklist.append(None)
      self.data[i] = weeklist
