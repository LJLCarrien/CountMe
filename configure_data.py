"""数据配置管理类"""
import json
from item_analysis_title import ColAnalysisTitleItem, RowAnalysisTitleItem
from item_title import TitleItem


class ConfigureData():
  """配置类"""
  def __init__(self, jsonfile_path):
    self.filepath = jsonfile_path
    self.jsondic = None
    self.secondtitle = {}

    self.defaulttitle_width = None
    self.defaulttitle_height = None

    self.default_fontname = None
    self.default_fontsize = None
    self.default_fontcolor = None
    self.default_bgcolor = None
    self.default_horalignment = None
    self.default_veralignment = None

    self.__openfile()

  def __dictitem2class(self, keyname: str, tmplist: list) -> list:
    '''json数组里的dict对象转换成对应的类对象'''
    resultlist = []
    for item in tmplist:
      if isinstance(item, dict):
        if keyname == 'title':
          titleitem = TitleItem()
          titleitem.set_data(self.jsondic, item)
          resultlist.append(titleitem)
        elif keyname == 'analysis_coltitle':
          coltitleitem = ColAnalysisTitleItem()
          coltitleitem.set_data(item)
          resultlist.append(coltitleitem)
        elif keyname == 'analysis_rowtitle':
          coltitleitem = RowAnalysisTitleItem()
          coltitleitem.set_data(item)
          resultlist.append(coltitleitem)
        else:
          print(f'{keyname}未支持')
      else:
        return tmplist
    return resultlist

  def __openfile(self):
    if self.filepath is None:
      return None
    with open(self.filepath, 'rb') as f:
      jsondic = json.load(f)
      self.jsondic = jsondic
      if isinstance(jsondic, dict):
        for key, value in jsondic.items():
          if isinstance(value, list):
            self.__dict__[key] = self.__dictitem2class(key, value)
          else:
            self.__dict__[key] = value

  def get_titlelist(self) -> list:
    '''获取一级菜单列表'''
    return self.title

  @staticmethod
  def get_value_ifnone_usedefault(defaultvalue, value):
    if value is not None:
      return value
    return defaultvalue

  def get_titletotal_line(self) -> int:
    '''获取标题所占行数（包括二级标记在内）'''
    result = 1
    sec_num = len(self.secondTitle)
    if sec_num >= 1:
      result = result + 1
    return result

  def get_title_height(self):
    '''行高'''
    return self.defaultTitleHeight

  def get_rowsum_dickeys(self) -> list:
    result = 'rowSumDicKey' in self.jsondic
    if result:
      return self.rowSumDicKey
    return None

  def get_colsum_dickeys(self) -> list:
    result = 'colSumDicKey' in self.jsondic
    if result:
      return self.colSumDicKey
    return None

  def get_weekcount_key(self) -> str:
    '''周合计'''
    return "weekCoust"

  def get_monthcount_key(self) -> str:
    '''月合计'''
    return "monthCoust"

  def get_format_dic(self) -> dict:
    '''获取fomat'''
    result = 'format' in self.jsondic
    if result:
      return self.format
    return None

  def get_excelpath(self) -> str:
    '''获取保存'''
    result = 'excelPath' in self.jsondic
    if result:
      return self.excelPath
    return None

  def get_isfreezemonth(self) -> bool:
    '''冻结行模式是月份'''
    return self.freezeLineMode['type'] == 'month'

  def get_isfreezeline(self) -> bool:
    '''冻结行模式是行数'''
    return self.freezeLineMode['type'] == 'line'

  def get_freezeline_detail(self) -> str:
    '''冻结行模式具体内容，月份时是1-12，行数是2到正无穷'''
    return self.freezeLineMode['detail']

  def get_analysis_rowtitle_list(self) -> list:
    '''分析结果-行菜单list'''
    return self.analysis_rowtitle

  def get_analysis_coltitle_list(self) -> list:
    '''分析结果-列菜单list'''
    return self.analysis_coltitle