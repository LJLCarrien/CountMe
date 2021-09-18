

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

##### 	全局默认：defaultFontColor

​		默认null时是黑色

##### 	局部默认：fontColor

##### 		一级标题：fontColor

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

##### 	全局默认：defaultBgColor

​		默认null时是黑色

##### 	局部默认：bgColor

##### 		一级标题：bgColor

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

针对列标题增加列，专门用于数字统计

##### 全局默认：暂无需求，不支持

##### 局部默认：暂无需求，不支持

##### 一级标题：fontColor

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

##### 	举例：dayMustCount

​		表示：日必须=餐饮+生活用品+交通+住房

```json
{
	"title": [
		{
			"name": "餐饮",
            "countCell": "合计",
			"dayMustCount": true
		},
		{
			"name": "生活用品",
            "countCell": "合计",
			"dayMustCount": true
		},
		{
			"name": "交通",
            "countCell": "合计",
			"dayMustCount": true
		},
		{
			"name": "住房",
            "countCell": "合计",
			"dayMustCount": true
		},
		{
			"name": "日必须",
			"rowSumDic": "dayMustCount",
		},
	],
	"rowSumDicKey": [
		"dayMustCount"
	],
}
```

##### 	举例：dayCount

​		表示：日合计=日必须+娱乐+购物

```json
{
	"title": [
		{
			"name": "日必须",
			"dayCount": true
		},
		{
			"name": "娱乐",
            "countCell": "合计",
			"dayCount": true
		},
		{
			"name": "购物",
            "countCell": "合计",
			"dayCount": true
		},
		{
			"name": "日合计",
			"rowSumDic": "dayCount",
		},
	],
	"rowSumDicKey": [
		"dayCount"
	],
}
```

#### 列求和

- ​	"colSumDicKey"-数组：求和列关键词


- ​	colSumDic"-属性：列求和结果关键字，一般是用来形容列求和的结果解释

  ​	列关键字做其他列的属性且为true时，表示求和此列。

##### 举例：weekCount、MonthCount

weekCount和MonthCount有自身逻辑，配置只是支持求和哪一列，求和结果在哪一列。

从配置上看，并没有像行求和那么智能（可以指定求和的元素列），但是目前也没有这样的需求。

```
{	
	"title": [
		{
			"name": "日合计",
			"weekCount": true,
			"MonthCount": true
		},
		{
			"name": "周合计",
			"fontColor": "white",
			"bgColor": "#b05574",
			"colSumDic": "weekCount"
		},
		{
			"name": "月合计",
			"fontColor": "white",
			"bgColor": "#dca09e",
			"colSumDic": "MonthCount"
		}
	],
	"colSumDicKey": [
		"weekCount",
		"MonthCount"
	],
}
```



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

        ​	关键字：dayMustCount

      - 日合计

        ​	关键字：dayCount

        ​	日合计=日必须+日消耗

    - [x] 列求和

      colSumDic标记求和结果的列位置，其value作为其他列的key且为true时，此时的其他列被标记为求和元素；colSumDicKey记录colSumDic的value值。

      关键字：colSumDic、colSumDicKey

      - 周合计

        ​	合计频率：从月开始到周日；从周一到周日；从周一到月结束；

        ​	关键字：weekCount

      - 月合计

        ​	关键字：monthCount

- #### 增加功能
  
  - [x] 支持配置excel文件名
  - [ ] 冻结表格指定行列功能，需要支持配置
  
- #### 代码优化

  - [x] 按功能组织成类

    - [x] 格式
    - [x] 配置
    - [ ] 行求和、列求和

  - [ ] 根据规范优化一次

    参考[Python风格规范](https://zh-google-styleguide.readthedocs.io/en/latest/google-python-styleguide/python_style_rules/#id1)

- #### 资料文档

  - [x] 更新README
  - [ ] 类结构关系图

### 第二阶段：数据存储

- [ ] todo:数据保存（json格式)

### 第三阶段：数据统计

- [ ] todo:折线图等,得出结论