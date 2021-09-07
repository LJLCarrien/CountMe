# CountMe
## 背景

​	统计我，数据化我，得出结论吧。

​	将平常的数据留下来，了解自己的消费习惯、消费结构，优化消费结构，提高生活水平，更加了解自己，基于已发生的事实反思，然后有针对性的执行，不断的优化自己的行为模式，看，这漫长的一生，我们最终提交的答卷。

## 环境

```bash
pip install xlsxwriter
```



## 配置解析 countMe_data.json 

### 配置：标题类别

​	在title里增加行标题类别

举例：增加日期列，name属性是日期。

```json
{
	"title": [
		{
			"name": "日期",
		},
	]
}
```

​	

### 配置：字体样式

#### 	列标题字体：titleFontName

```json
{
	"titleFontName": "宋体",
}
```

### 配置：字体大小

#### 	列标题字体大小：titleFontSize

```json
{
	"titleFontSize": 13
}
```

#### 	列二级标题字体大小：secTitleFontSize

```json
{
	"secTitleFontSize": 11,
}
```



### 配置：宽高

#### 列宽

##### 	默认配置：defaultTitleWidth

##### 	个别配置：width

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

##### 	默认配置：defaultTitleHeight

```json
{
	"defaultTitleHeight": 17,
}
```



### 配置：颜色

#### 列标题字体颜色 fontColor 

- fontColor 默认是黑色，没有该属性就是黑色；

#### 列标题背景颜色 bgColor

- bgColor 默认是无背景颜色，没有该属性即无背景颜色。

#### 举例

- ​	“周末”：白色字体，无背景。（护眼模式可以看到该标题）
- ​	“ToDo”：红色字体，无背景。
- ​	“餐饮”：白色字体，背景颜色为#fe4365。

```json
{
	"title": [
		{
			"name": "周目",
			"fontColor": "white"
		},
		{
			"name": "ToDo",
			"fontColor": "red"
		},
		{
			"name": "餐饮",
			"fontColor": "white",
			"bgColor": "#fe4365",
		},
	],
}
```

### 配置：合计 countCell

该列增加合计列，专门用于数字统计

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



### 配置：行求和

- "rowSumDicKey"-数组：列举所有行求和关键词
- "rowSumDic"-属性：行求和结果关键字，一般是用来形容该列求和的结果解释
  对应的关键字做其他列的属性且为true时，表示此列为对应属性行求和时的元素之一。

#### 举例：dayMustCount

​	表示日必须=餐饮+生活用品+交通+住房

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

#### 举例：dayCount

​	表示日合计=日必须+娱乐+购物

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





## 计划

- 第一阶段：表格创建

  - 基本框架
    - [x] 配置列宽行高
    - [x] 列标题
      - [x] 配置 类别
      - [x] 配置 字体样式、字体大小、字体颜色、背景颜色
    - [x] 行日期
      - [x] 字体样式、字体大小、字体颜色、背景颜色（代码写死/默认，无配置功能）
    - [x] 行周目
      - [x] 字体样式、字体大小、字体颜色、背景颜色（代码写死/默认，无配置功能）
    - [x] 合计
      - [x] 限制格式，2位小数的数字
      - [x] 字体样式、字体大小、字体颜色、背景颜色（代码写死/默认，无配置功能）

  - 基本逻辑
    - [x] 按日期逐行显示
    - [ ] 一个表格生成全年数据

  - 基本功能
    - [x] 行求和（日必须、日合计）
    - [x] 列求和（周合计、月合计）
  - 代码优化（部分格式改为支持配置）
    - [ ] 支持配置excel文件名、表格名
    - [ ] 按功能组织成类
      - [ ] 格式类
      - [ ] 读取文件数据为对象

- 第二阶段：数据存储
  - [ ] todo:数据保存（json格式)

- 第三阶段：数据统计
  - [ ] todo:折线图等,得出结论