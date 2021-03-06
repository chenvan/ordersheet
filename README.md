## 水分汇总

- 回潮时间中年份改回2020年
- 时间加减用 TEXT 公式: TEXT(时间相减, "[h]:mm")



## vbaDeveloper的使用

<https://github.com/hilkoc/vbaDeveloper>

手动打开vbaDeveloper.xlam, 就会出现加载项, 然后导出

**excel文件名是中文就无法导出**



##  VBA-JSON的使用

<https://github.com/VBA-tools/VBA-JSON>



## HDT烘丝入口水分

跨文件也可以引用



## 时间转换

输入3个或4个数字, 然后转换为时间, 通过日期栏的日期, 把日期也填补上. 

如果是加班, 则自己手动修改

或者需要输入大于2400的数字(?)



## 语音提醒

### 锚点

锚点1: 开始时间(经常容易忘记)

- 开始时间容易忘记写,需要提醒
- 提醒的是这一批次烟**开始后**需要注意的东西, 如查看入柜状态,注意变异系数, 注意精度



锚点2: 结束时间, 提醒下一批要写开始时间,

- 提醒这一批次**结束后**需要注意的地方
- 提醒下一批次**开始前**需要检查的地方
- 提醒要写下一批次的开始时间



每批烟需要有三处提醒内容, 分别是**开机前, 开机后, 结束后**



### 第一批提醒

第一批开机前需要的提醒



## 语音提醒函数注意点

先检查日期

1. 日期栏是今天或者空白
   - 现在的时间没有超过deadLineOffset + 本来要播报的时间就进行播报
2. 日期栏过期
   - 不进行提醒


​	

## 语音提醒取消

时间栏改变就会取消原来的语音提醒



## 状态栏文字提醒

类似FIFO的列表



## 解决无法立即更新工作薄中的部分链接

HDT段的表先改为不保护的状态(未完成)



## 时间预估

预估收工时间



## 使用JSON记录语音数据

***原则: 数据不要分散放置, 因为一旦要修改就需要同时修改多个地方, 容易漏掉***

牌号相关的数据, 如加料比例, 加香比例,主叶丝秤的流量, 梗丝,膨丝,残丝的掺配比例等等, 放到excel表中

语音提醒放到JSON文件中(其中柜的分类工作直接在代码中完成)



### 文件结构
```json
{
    "回潮": {

        "第一批": [Node, ...],

        "开始前": [Node, ...],

        "开始后": [Node, ...],

        "结束后": [Node, ...]
    },

    "加料": {

        "第一批": [Node, ...],

        "开始前": [Node, ...],

        "开始后": [Node, ...],

        "结束后": [Node, ...]

    },

    "切烘加香": {

        "第一批": [Node, ...],

        "开始前": [Node, ...],

        "开始后": [Node, ...],

        "结束后": [Node, ...]

    },

    "贮叶柜": {

        "叶柜":  [Node, ...],

        "HDT":  [Node, ...]

    },

    "贮丝柜": {

        "南区AF":  [Node, ...],

        "南区GM":  [Node, ...],

        "南区NT":  [Node, ...],

        "木箱AC":  [Node, ...],

        "北区AD":  [Node, ...],

        "北区EH":  [Node, ...]

    }
}
```


### Node的结构
```json
{

    "sOffsetTime": int,

    "aOffsetTime": int,

    "deadLineOffset": int,
    
    "filter": array,
    
    "isSuccessive": boolen,
    
    "mode": string,

    "redirectList": array, 

    "content": string 

}
```

deadLineOffset +本来应播报的时间就是最后需要播报该语音的时间

sOffsetTime与aOffsetTime两者只会有一

​	sOffsetTime是固定的偏移时间

​	aOffsetTime则需要经过主叶丝秤流量进行变换(输送带的速度没有改变, 主要是填满主叶丝暂存柜需要的时间不同)

触发条件: 需要先检查触发条件是否存在, 如果不存在就表示该语音普遍适用

​	filter是对牌号的筛选, 只有在filter中的牌号才会触发语音

​	isSuccessive牌号是否连续. 如果是开始前,开始后的语音, isSuccessive指的是与上一批是否同牌号, 如果是结束后, 则指的是与下一批是否为同牌号. 有些语音需要在牌号连续的时候播报, 有些则需要在不是连续的牌号时播报

​	mode进柜方式.  "half", "full"

redirectList, 与content二者只会有一

​	存在redirectList. 用redirectList, 牌号找到需要的参数, 用参数生成新的内容

​	存在content. 直接使用



### 其他注意点

JSON 文件的编码需要是 ANSI, 否则中文会变为乱码且无法正确解析



## 如何改变偏移时间

***使用双喜(经典)的偏移时间作为基准***

主叶丝秤之后的语音偏移时间才需要改变



## Error的处理

util里的函数需要有Error的处理, 这样可以方便的定位错误的原因



## 可能存在的Bug

晚上加班可能会让时间函数判断不用提醒



## 条件格式

### 高亮选中行和列

条件公式范围内的格子: 使用 row 和 column 函数

选中的格子: 使用 cell 函数



#### 序号为第5, 9, 则为true

列: 选中格所在行的第二列是5或者9, 那么列数与选中格一样的格子将为true

```
AND(CELL("col") = COLUMN(), OR(INDIRECT(ADDRESS(CELL("row"), 2)) = 5, INDIRECT(ADDRESS(CELL("row"), 2)) = 9))
```
行: 选中格所在行的第二列是5或者9, 那么行数与选中格一样的格子将为true

```
AND(CELL("row") = ROW(), OR(INDIRECT(ADDRESS(CELL("row"), 2)) = 5, INDIRECT(ADDRESS(CELL("row"), 2)) = 9))
```
#### 序号不是第5, 9, 则为true

列
```
AND(CELL("col")=COLUMN(), INDIRECT(ADDRESS(CELL("row"), 2)) <> 5, INDIRECT(ADDRESS(CELL("row"), 2)) <> 9)
```
行
```
AND(CELL("row")=ROW(), AND(INDIRECT(ADDRESS(CELL("row"), 2)) <> 5, INDIRECT(ADDRESS(CELL("row"), 2)) <> 9))
```



### 序号为1的, 则为true
```
$B2=1 
```

由于采用的格式上横线是红色,但是下横线是黑色. 所以如果连续两个序号为1的中间线不会变成红色(在弄一个格式)

### 加料精度超标, 则为true
```
ABS($L3-$K3)/$K3 > 0.01
```

### 自己序号为1, 上一行的序号也为1, 则为true(未完成)
```
AND(INDIRECT(ADDRESS(ROW(), 2)) = 1, INDIRECT(ADDRESS(ROW()-1, 2)) = 1)
```



## 解决其他表插入列后会令水分汇总表数据错乱的问题

vlookup写对了可以自动应对插入列的问题

首先第二个参数"寻找的范围"可以自动更新

第三个参数使用 colum(表名!单元格1) - colum(表名!单元格2) + 1 也能进行自动更新



## 烘丝料头温度的升温过程
1. 起始点温度知道
2. 最终温度在变化(变化不大)
3. 升温时间已知
3. 温度惯性较大, 升温过程需要先快后慢

## 回掺



## 水分仪修改



## NEXT

1. offset时间精确到秒
2. ~~状态栏文字提醒加上时间~~(时间格式不一样)
3. 提醒的过期逻辑需要重新思考
4. 语音类型分类
5. 搞一个烦人语音模式?
6. 主页丝秤流量
7. 增加半柜的记录
8. 增加测算烟包平均水分
9. 水分标红的地方需要修改

