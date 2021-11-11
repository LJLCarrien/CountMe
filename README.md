# CountMe

## 背景

​	统计我，数据化我，得出结论吧。

​	将平常的数据留下来，了解自己的消费习惯，优化消费结构，理性消费，提高生活水平，更加了解自己，基于事实反思，然后有针对性得执行，不断得优化自己的行为模式，在不可控的人生里，慢慢地重新掌握自己的生活。做自己的主人，不亦步亦趋，不骄不躁，平静而有力量，在自己能力范围内，给自己最棒的生活体验，有干劲的时候尝试突破自己，累了有底气好好休息，周而复始，不断往复。看，这漫长的一生，我们最终提交的答卷，检查数遍，死而无憾。

## 环境

```bash
pip install xlsxwriter
```



## 配置解释 config.json 

### 结构介绍

​	格式优先级：全局默认 < 局部默认 < 特定

#### 全局默认格式

​	关键字格式为：defaultxxxxxx

```json
{
	"defaultTitleWidth": 11,
	"defaultTitleHeight": 17,
	
	"defaultBold": false,
	"defaultHorAlignment": "center",
	"defaultVerAlignment": "vcenter",
	"defaultFontSize": 11,
	"defaultFontName": "宋体",
	"defaultFontColor": null,
	"defaultBgColor": null,
}
```

#### 局部默认格式

​	format对象里，例如"title"指的是一级标题默认格式，”secTitle“：指的是二级标题默认格式

```json
{
	"format": {
		"title": {
			"fontSize": 13,
			"fontName":"微软雅黑",
			"bold": true
		},
		"secTitle":{
			"fontName":"微软雅黑",
			"bold": true
		}
	}
}
```

#### 一级标题

​	在title[]里增加标题对象，举例：增加日期列，name属性是日期。

```json
{
	"title": [
		{
			"name": "日期",
		},
	]
}
```



### 宽高

#### 列宽

##### 	全局默认：defaultTitleWidth

##### 	局部默认：无意义，不支持

##### 		一级标题：width

```json
{
	"defaultTitleWidth": 11,
	"title": [
		{
			"name": "ToDo",
			"width": 20,
		},
	]
}
```

#### 行高

##### 	全局默认：defaultTitleHeight

##### 	局部默认：无意义，不支持

##### 		一级标题：无意义，不支持

​			标题是列标题，配置行高无意义。

```json
{
	"defaultTitleHeight": 17,
}
```



### 粗体

#### 全局默认：defaultBold

#### 局部默认：bold

#### 一级标题：丑拒，不支持

​		因为个人审美上觉得：要么标题就是统一加或不加粗会比较好看，实在是没有特意加上这个功能的动力。

```json
{
	"defaultBold": false,
	"format": {
		"title": {
			"bold": true
		},
	}
}
```

#### 

### 对齐（水平、垂直）

#### 值参考

​	[alignment](https://xlsxwriter.readthedocs.io/format.html?highlight=center#format-set-align)

| Horizontal alignment |
| :------------------- |
| left                 |
| center               |
| right                |
| fill                 |
| justify              |
| center_across        |
| distributed          |

| Vertical alignment |
| :----------------- |
| top                |
| vcenter            |
| bottom             |
| vjustify           |
| vdistributed       |

#### 全局默认：defaultHorAlignment、defaultVerAlignment

#### 局部默认：horAlignment、verAlignment

#### 一级标题：丑拒，不支持

​		因为个人审美上觉得：标题统一对齐会比较好看，实在是没有特意加上这个功能的动力。

```json
{
	"defaultHorAlignment": "center",
	"defaultVerAlignment": "vcenter",
	"format": {
		"weekDay": {
			"horAlignment": "left",
			"verAlignment": "vcenter"
		},
	}
}
```

#### 

### 字体大小

#### 全局默认：defaultFontSize

#### 局部默认：fontSize

​	举例：日期，结果为13

​	"defaultFontSize": 11-》"fontSize": 13

#### 	一级标题：fontSize

​	举例：ToDo，结果为18

​	"defaultFontSize": 11-》"fontSize": 13 -》 "fontSize": 18

```json
{
	"defaultFontSize": 11,
	"format": {
		"title": {
			"fontSize": 13,
		}		
	},
	"title": [
        {
			"name": "日期",
		},
		{
			"name": "ToDo",
			"fontSize": 18,
		}
	]
}
```



### 字体样式

优先级：全局默认 < 局部默认 < 特定

#### 全局默认：defaultFontName

#### 局部默认：fontName

- ​	举例：周目，结果为幼圆

​	"defaultFontName": "宋体"-》"fontName":"幼圆"

- ​	举例：日期，结果为微软雅黑

​	"defaultFontName": "宋体"-》"fontName":"微软雅黑"

#### 	一级标题：fontName

​	举例：ToDo，结果为方正姚体

​	"defaultFontName": "宋体"-》"fontName":"微软雅黑" -》"fontName":"方正姚体"

```json
{
	"defaultFontName": "宋体",
	"format": {
		"title": {
			"fontName":"微软雅黑",
		},
		"weekDay": {
			"fontName":"幼圆"
		}
	},
	"title": [
        {
			"name": "日期",
		},
		{
			"name": "ToDo",
			"fontName":"方正姚体"
		},		
	],
}
```

### 

### 颜色

#### 值参考

[working with colors](https://xlsxwriter.readthedocs.io/working_with_colors.html#colors)

| Color name | RGB color code |
| :--------- | :------------- |
| black      | `#000000`      |
| blue       | `#0000FF`      |
| brown      | `#800000`      |
| cyan       | `#00FFFF`      |
| gray       | `#808080`      |
| green      | `#008000`      |
| lime       | `#00FF00`      |
| magenta    | `#FF00FF`      |
| navy       | `#000080`      |
| orange     | `#FF6600`      |
| pink       | `#FF00FF`      |
| purple     | `#800080`      |
| red        | `#FF0000`      |
| silver     | `#C0C0C0`      |
| white      | `#FFFFFF`      |
| yellow     | `#FFFF00`      |

#### 字体颜色

##### 		全局默认：defaultFontColor

​		默认null时是黑色

##### 		局部默认：fontColor

##### 			一级标题：fontColor

```json
{
	"defaultFontColor": null,
	"format": {
		"countNumCell": {
			"fontColor": "red"
		},
		"monthEnd": {
			"fontColor": "#a12f2f"
		}
	},
	"title": [
        {
			"name": "周目",
			"fontColor": "white"
		},
		{
			"name": "ToDo",
			"fontColor": "red"
		}	
	],
}
```



#### 背景颜色

##### 		全局默认：defaultBgColor

​		默认null时是黑色

##### 		局部默认：bgColor

##### 			一级标题：bgColor

```json
{
	"defaultBgColor": null,
	"format": {
		"monthEnd": {
			"bgColor": "#dca09e"
		}
	},
	"title": [
        {
			"name": "餐饮",
			"bgColor": "#fe4365"
		},
	],
}
```



### 合计

​	针对列标题增加列，专门用于数字统计

#### 全局默认：暂无需求，不支持

#### 局部默认：暂无需求，不支持

#### 一级标题：fontColor

```
{
	"title": [
		{
			"name": "餐饮",
			"countCell": "合计",
		},
		{
			"name": "生活用品",
			"countCell": "合计",
		},
		{
			"name": "交通",
			"countCell": "合计",
		},
		{
			"name": "住房",
			"countCell": "合计",
		}
	]
}
```

### 表格冻结

​	功能参考：[freeze_panes](https://xlsxwriter.readthedocs.io/worksheet.html?highlight=freeze#freeze_panes)

#### 指定冻结列位置

##### 一级标题：freeze

​	默认标题完整显示（即上固定区域和下滚动区域**不支持**配置），**仅支持**配置冻结列位置，即可以影响左固定区域和右滚动区域

​	"freeze":false 时不冻结表格；出现多个"freeze":true，以最后一个为准，因此不建议标题列处出现多个freeze。

###### 	**举例：单个freeze**

​	指定Todo列为右滚动区域第一列

```
{
	"title": [
		{
			"name": "ToDo",
			"freeze":true
		}
	]
}
```

###### 	**举例：多个freeze**

​	指定Todo和HaveDone时，结果为HaveDone

```
{
	"title": [
		{
			"name": "ToDo",
			"freeze":true
		},
		{
			"name": "HaveDone",
			"freeze":true
		}
	]
}
```



#### 指定滚动行

用于表格冻结后，右滚动区域显示的首行

以下两个例子效果一致

##### 	举例：指定月份模式

​	右滚动区域显示指定月份为首行

```json
{
	"freezeLineMode":{
		"type":"month",
		"detail":"1"
	}
}
```

##### 	举例：指定行模式

​	右滚动区域显示指定表格行数为首行

```
{
	"freezeLineMode":{
		"type":"line",
		"detail":"10"
	}
}
```



### 复合功能

#### 二级标题

​	secondSort：二级菜单种类名，二级菜单名为value值。

​	在 secondTitle  里寻找与value名一致的数组，若数组长度大于零，意味着有二级菜单。

##### 	举例：餐饮 

​	secondSort值为“diet"（其他英文也可），secondTitle里添加对应的diet数组

```json
{
	
	"title": [
		{
			"name": "餐饮",
			"secondSort": "diet",
		},
	],
	"secondTitle": {
		"diet": [
			"早餐",
			"午餐",
			"晚餐",
			"宵夜",
			"饮食饮料",
			"外卖"
		]
	}
}
```

##### 举例：购物

购物secondSort值为“Shopping"，在secondTitle里添加Shopping数组

```json
{	
	"title": [
		{
			"name": "购物",
			"secondSort": "Shopping",
		},
	],
	"secondTitle": {
		"Shopping":[
			"上衣",
			"裤子"
		]
	}
}
```



#### 行求和

- "rowSumDicKey"-数组：列举所有行求和关键词
- "rowSumDic"-属性：行求和结果关键字，一般是用于行求和的结果解释
  对应的关键字做其他列的属性且为true时，表示此列为对应属性行求和时的元素之一。

##### 	举例：dayMustCoust

​		表示：日必须=餐饮+生活用品+交通+住房

```json
{
	"title": [
		{
			"name": "餐饮",
            "countCell": "合计",
			"dayMustCoust": true
		},
		{
			"name": "生活用品",
            "countCell": "合计",
			"dayMustCoust": true
		},
		{
			"name": "交通",
            "countCell": "合计",
			"dayMustCoust": true
		},
		{
			"name": "住房",
            "countCell": "合计",
			"dayMustCoust": true
		},
		{
			"name": "日必须",
			"rowSumDic": "dayMustCoust",
		},
	],
	"rowSumDicKey": [
		"dayMustCoust"
	],
}
```

##### 	举例：dayCoust

​		表示：日合计=日必须+娱乐+购物

```json
{
	"title": [
		{
			"name": "日必须",
			"dayCoust": true
		},
		{
			"name": "娱乐",
            "countCell": "合计",
			"dayCoust": true
		},
		{
			"name": "购物",
            "countCell": "合计",
			"dayCoust": true
		},
		{
			"name": "日合计",
			"rowSumDic": "dayCoust",
		},
	],
	"rowSumDicKey": [
		"dayCoust"
	],
}
```

#### 列求和

- ​	"colSumDicKey"-数组：求和列关键词


- ​	colSumDic"-属性：列求和结果关键字，一般是用来形容列求和的结果解释

  ​	列关键字做其他列的属性且为true时，表示求和此列。

##### 举例：weekCoust、MonthCount

weekCount和MonthCount有自身逻辑，配置只是支持求和哪一列，求和结果在哪一列。

从配置上看，并没有像行求和那么智能（可以指定求和的元素列），但是目前也没有这样的需求。

```
{	
	"title": [
		{
			"name": "日合计",
			"weekCoust": true,
			"MonthCount": true
		},
		{
			"name": "周合计",
			"fontColor": "white",
			"bgColor": "#b05574",
			"colSumDic": "weekCoust"
		},
		{
			"name": "月合计",
			"fontColor": "white",
			"bgColor": "#dca09e",
			"colSumDic": "MonthCount"
		}
	],
	"colSumDicKey": [
		"weekCoust",
		"MonthCount"
	],
}
```

## 存储数据

参考：save_data.json

"emptyTemplate":为保存数据对象类生成对应属性用的，千万不要删除。

```json
{
    "emptyTemplate": {
        "todo": "",
        "havedone": "",
        "diet": {
            "breakfast": "",
            "lunch": "",
            "dinner": "",
            "nightsnack": "",
            "snacksdrinks": "",
            "takeoutfood": "",
            "countcell": null
        },
        "dailyuse": {
            "des": "",
            "countcell": null
        },
        "traffic": {
            "des": "",
            "countcell": null
        },
        "housing": {
            "des": "",
            "countcell": null
        },
        "daymustcoust": {
            "countcell": 0
        },
        "entertainment": {
            "des": "",
            "countcell": null
        },
        "shopping": {
            "des": "",
            "countcell": null
        },
        "daycoust": {
            "countcell": 0
        }
    }
}
```

测试数据文件（具体参考：save_data_for_test.json）



## 计划

### 第一阶段：表格创建

- #### 基本逻辑

  - [x] 年份为表格名
  - [x] 全年数据生成在一个表格

- #### 基本框架

  - [x] 显示类格式配置

    格式优先级：全局默认格式 < 局部默认格式 < 特定格式

    - [x] 全局默认格式（“default...”)

      - 行宽、列高

        defaultTitleWidth、defaultTitleHeight

      - 对齐方式：水平、垂直

        defaultHorAlignment、defaultVerAlignment

      - 加粗

        defaultBold

      - 字体大小、字体样式、字体颜色、背景颜色

        defaultFontSize、defaultFontName、defaultFontColor、defaultBgColor

    - [x] 局部默认格式 （“format”)

      - 单元格数字格式 

        numFormat

      - 对齐方式：水平、垂直

        horAlignment、verAlignment

      - 加粗 

        bold

      - 字体大小、字体样式、字体颜色、背景颜色

        fontSize、fontName、fontColor、bgColor

    - [x] 一级标题配置（“title”:[]）

      - 行宽、列高 

        width、height

      - 字体大小、字体样式、字体颜色、背景颜色

        fontSize、fontName、fontColor、bgColor

  - [x] 基础功能

    - [x] 一级标题

      ​	无二级标题的一级标题需要与第二行合并。

      ​	格式优先级：全局默认格式 < **一级标题**局部默认格式 < **一级标题**配置（title[])

      ​	关键字：name

    - [x] 二级标题

      ​	格式优先级：全局默认格式 < **二级标题**局部默认格式 < **一级标题**配置（title[])

      ​	关键字：secondSort、secondTitle

    - [x] 日期、周目

      按日期逐行显示，第一列日期，第二列周目

      日期格式优先级：全局默认格式 < **日期**局部默认格式（date）

      周目格式优先级：全局默认格式 < **周目**局部默认格式（weekDay）

    - [x] 合计

      此为纯数字，用于统计数据。

      格式优先级：全局默认格式 < **合计**局部默认格式（countNumCell）

      关键字：countCell

    - [x] 行求和

      rowSumDic标记求和结果的列位置，其value作为其他列的key且为true时，此时的其他列被标记为求和元素；rowSumDicKey 记录rowSumDic的value值。

      关键字：rowSumDic、rowSumDicKey

      - 日必须

        ​	关键字：dayMustCoust

      - 日合计

        ​	关键字：dayCoust

        ​	日合计=日必须+日消耗

    - [x] 列求和

      colSumDic标记求和结果的列位置，其value作为其他列的key且为true时，此时的其他列被标记为求和元素；colSumDicKey记录colSumDic的value值。

      关键字：colSumDic、colSumDicKey

      - 周合计

        ​	合计频率：从月开始到周日；从周一到周日；从周一到月结束；

        ​	关键字：weekCoust

      - 月合计

        ​	关键字：monthCoust

- #### 增加功能
  
  - [x] 支持配置excel文件名
  - [x] 冻结表格，需要支持配置
    - [x] 指定列功能
    - [x] 指定滚动行功能（指定月份；指定行数）
  
- #### 代码优化

  - [x] 按功能组织成类

    - [x] 格式
    - [x] 配置
    - [x] 行求和、列求和

  - [x] 根据规范优化一次

    参考[Python风格规范](https://zh-google-styleguide.readthedocs.io/en/latest/google-python-styleguide/python_style_rules/#id1)
    
    使用代码检查工具 pylint
    
    - 常量大写
    - 类总是使用驼峰式命名，即单词**首字母大写**其余字母小写
    - 实例函数名，全小写，下划线分割单词
    
    参考：[PEP 8 -- Python 代码风格指南](https://www.python.org/dev/peps/pep-0008/)
    
    使用自动格式化工具yapf

- #### 资料文档

  - [x] 更新README
  - [ ] 类结构关系图

### 第二阶段：数据存储

目前设定是：python读取json数据，生成excel文件。

由另一个客户端/应用操作后，保存生成json数据。（这是后面的事情了，不在这里考虑）

- [x] 数据保存（json格式)
  暂时是手动生成的
- [x] 读取save_data.json，写入xlsx

### 第三阶段：数据统计

- [ ] todo:折线图等,得出结论

  

## Q&A 

### 为什么要做这个项目直接操作excel不好吗？

​	可以是可以，但是光操作excel没意思。我对格式有要求，太丑我看不下去，但是如果每月、每年都要不时搞格式，会让我懒得开始。每每在做重复的格式操作的瞬间，我都在心里吐槽“害，好无聊啊，我只想管记录，什么时候才可以用代码生成啊”。

​	用代码生成的好处是，只需要留有一份数据（此处指json文件），就可以生成格式符合自己预期的excel文件。我已经纯用excel记录日常好几年了，很多时候我都好想对比一下自己几年来的数据，但是我又很懒，光是看数据我头又很大，几个excel之间操作我也觉得无聊又麻烦，如果每年都要这样操作一下，光是想想都开始劝退，可是这个对比的想法又不断的生根发芽，如果有人可以用帮我将几年的数据用图表展示出来，一年里每个月的消费用图表统计下，就好了，所以那个人就是我了。

​	我希望excel只是作为展示最终结果，操作时使用更加美观简洁的界面专注在记录就好。从这个操作界面到excel展示要做的事情就全部交给python了。虽然我有考虑过要不要做excel内容由python直接生成json，但是我个人想摆脱excel操作才做的这些事情，目前没有强而有力的理由，不做的理由倒是蛮充分的，所以还是不实现这个功能了。

