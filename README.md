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



### 其他

加料批次的转烟时间, 同批次的转烟时间要20分钟左右, 不同批次的要分钟30左右

切烘加香第四批, 第八批需清扫



## 语音提醒函数注意点

先检查日期

1. 日期栏是今天或者空白
   - 时间已过, 如果isForceBroadcast 是 true , 立即进行提醒  
2. 日期栏过期
   - 不进行提醒



##状态栏文字提醒

类似FIFO的列表



## 解决无法立即更新工作薄中的部分链接

HDT段的表先改为不保护的状态(未完成)



## 使用JSON记录语音数据

***原则: 数据不要分散放置, 不要同时修改多个地方***

牌号相关的数据, 如加料比例, 加香比例,主叶丝秤的流量, 梗丝,膨丝,残丝的掺配比例等等, 放到excel表中

语音提醒放到JSON文件中(其中柜的分类工作直接在代码中完成)



###文件结构
```json
{
    回潮: {

        第一批: [Node, ...],

        开始前: [Node, ...],

        开始后: [Node, ...],

        结束后: [Node, ...]

    },

    加料: {

        第一批: [Node, ...]

        开始前: [Node, ...],

        开始后: [Node, ...],

        结束后: [Node, ...]

    },

    切烘: {

        第一批: [Node, ...]

        开始前: [Node, ...],

        开始后: [Node, ...],

        结束后: [Node, ...]

    },

    贮叶柜: {

        叶柜:  [Node, ...],

        HDT:  [Node, ...]

    },

    贮丝柜: {

        南区AF:  [Node, ...],

        南区GM:  [Node, ...],

        南区NT:  [Node, ...],

        木箱AC:  [Node, ...],

        北区AD:  [Node, ...],

        北区EH:  [Node, ...]

    }
}
```


###Node的结构
```json
{

    sOffsetTime*: int,

    aOffsetTime*: int,

    isForceBroadcast: boolen,

    redirect*: string, 

    hdt*: string

    content*: string 

}
```

isForceBroadcast 是 true, 则超时也需要播报该语音

sOffsetTime与aOffsetTime两者只会有一

sOffsetTime是固定的偏移时间

aOffsetTime则需要经过主叶丝秤流量进行变换(输送带的速度没有改变, 主要是填满主叶丝暂存柜需要的时间不同)

redirect, hdt与content三者只会有一

存在redirect. 用redirect, 牌号找到需要的参数, 用参数生成新的内容

存在hdtContent. 用牌号检查是否需要回掺HDT, 需要回掺则使用hdtContent

存在content. 直接使用



### 其他注意点

JSON 文件的编码需要是 ANSI, 否则中文会变为乱码且无法正确解析



##如何改变偏移时间

***使用双喜(经典)的偏移时间作为基准***

主叶丝秤之后的语音偏移时间才需要改变

adjustOffsetTime = offsetTime * 6250 / 主叶丝秤流量 



##Error的处理



##可能存在的Bug

晚上加班可能会让时间函数判断不用提醒






