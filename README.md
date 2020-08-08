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
   - 如果时间没超过10分钟, 那么进行过期提醒(立即语音)
   - 超过10分钟不再提醒(进行提醒已经无作用)
2. 日期栏过期
   - 不进行提醒



## 状态栏文字提醒

类似FIFO的列表



## 解决无法立即更新工作薄中的部分链接

HDT段的表先改为不保护的状态(未完成)



## 使用JSON记录语音数据

### default 语音文件, 共同的提醒项目和开机前提醒

- 第一批
- 回潮段
- 加料段
- 切烘段



### 各牌号语音文件, 牌号相关的语音还有牌号的各种参数

牌号

- 第一批
- 回潮段
- 加料段
- 切烘段



### 柜的语音文件

柜号



### 其他注意点

JSON 文件的编码需要是 ANSI, 否则中文会变为乱码且无法正确解析



## 可能存在的Bug

晚上加班可能会让时间函数判断不用提醒





