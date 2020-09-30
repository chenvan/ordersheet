## 水分汇总

- 回潮时间中年份改回2020年
- 时间加减用 TEXT 公式: TEXT(时间相减, "[h]:mm")



## vbaDeveloper的使用

<https://github.com/hilkoc/vbaDeveloper>

手动打开vbaDeveloper.xlam, 就会出现加载项, 然后导出



##  VBA-JSON的使用

<https://github.com/VBA-tools/VBA-JSON>



## HDT烘丝入口水分

96线同步复制HDT数据(未完成)

- 从96的HDT表中遍历经典HDT, 如果64加料表有这张工单, 然后64HDT表没有这张工单, 就进行append

  

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



### 通用提醒

与机器相关的提醒



### 与牌号相关的提醒

牌号的工艺参数与其他设置



### 柜的提醒

不同的柜不同的延时



### 其他

加料批次的转烟时间, 同批次的转烟时间要20分钟左右, 不同批次的要分钟30左右

切烘加香第四批, 第八批需清扫



## 语音提醒函数注意点

先检查日期

1. 日期栏是今天或者空白
   - 时间已过, 如果isForceBroadcast 是 true , 立即进行提醒  
2. 日期栏过期
   - 不进行提醒



## 状态栏文字提醒

类似FIFO的列表



## 解决无法立即更新工作薄中的部分链接

HDT段的表先改为不保护的状态(未完成)



## 使用JSON记录语音数据

***原则: 数据不要分散放置, 不要同时修改多个地方***

牌号相关的数据, 如加料比例, 加香比例,主叶丝秤的流量, 梗丝,膨丝,残丝的掺配比例等等, 放到excel表中

语音提醒放到JSON文件中(其中柜的分类工作直接在代码中完成)



####文件结构

回潮: {

​		第一批: [Node, ...],

},

​		开始前: [Node, ...],

​		开始后: [Node, ...],

​		结束后: [Node, ...]

},

加料: {

​		第一批: [Node, ...]

​		开始前: [Node, ...],

​		开始后: [Node, ...],

​		结束后: [Node, ...]

},

切烘: {

​		第一批: [Node, ...]

​		开始前: [Node, ...],

​		开始后: [Node, ...],

​		结束后: [Node, ...]

},

贮叶柜: {

​		叶柜:  [Node, ...],

​		HDT:  [Node, ...]

},

贮丝柜: {

​		南区AF:  [Node, ...],

​		南区GM:  [Node, ...],

​		南区NT:  [Node, ...],

​		木箱AC:  [Node, ...],

​		北区AD:  [Node, ...],

​		北区EH:  [Node, ...]

}



####Node的结构

{

​		offsetTime: int,

​		isForceBroadcast: boolen,

​		isRedirect: boolen,

​        redirectIndex*: string, (isRedirect 是 true 时, 才有 redirectIndex)

​		content*: string (isRedirect 是 false 时, 肯定有 content, 是 true 时则不一定)

}

isForceBroadcast 是 true, 则超时也需要播报该语音

isRedirect 是 true, 用牌号 和 redirectIndex 作为索引检索数据, 如果索引成功, 则看是否有 content 属性, 有的话播报 content(解决掺HDT的播报), 没有的话用 牌号 + redirectIndex + 检索的数据 作为新的 content 

如果是切烘段, delay时间还需要根据主叶丝秤流量, 膨丝流量, 梗丝流量进行变换



### 其他注意点

JSON 文件的编码需要是 ANSI, 否则中文会变为乱码且无法正确解析



## 可能存在的Bug

晚上加班可能会让时间函数判断不用提醒











