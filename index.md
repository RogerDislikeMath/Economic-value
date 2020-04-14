# 经济模拟器
## 一、前言  
### 经济模拟器集成了游戏中大部分数值相关的经济系统，包括如下几个部分，以下针对Sheet页签进行说明。
	  
      1.Value-经济模拟器首页指引。
      2.Vow-抽奖模拟器(带按钮)，可根据调整充值金额&产出概率&分解装备进行真实模拟。
      3.装备升星消耗-装备升星需要消耗的材料&金币数量。
      4.Equipenhance-装备强化需要消耗的材料&金币数量。
      5.游戏礼包-游戏设计的各种礼包，包含价格价值折扣等。
      6.商城物品价值及分类-游戏中大部分物品&资源的对应价值。(价值≠价格≠钻石，但有对应关系)
      7.DropPlan-关卡掉落规划(粗略版)，自动化填表基础。
      8.ActualDrop-使用Excel自带公式进行的转换，将实际掉落需求分布从DropPlan转化到该页中。
      9.DropTable-代码最终生成的掉落片段均在这页中。
## 二、整体经济循环的设计
      
### 为了逻辑清楚，我们先从材料层面说明。(不包含付费礼包)
      
      1.装备的产出：抽卡系统。
      2.木、石、铁的产出：前期依靠关卡，中后期依靠家园"木石铁三合一建筑"。
      3.装备升星材料：60%在地牢，极少量关卡，10%在竞技场商城，白、绿装备大量产出与任务&活动等。
      4.装备强化材料：大量产出于关卡，少量产出于任务&活动。
      5.金币：关卡，任务&活动，竞技场商店，地牢，金矿。
      6.竞技场勋章：竞技场胜利，赛季奖励。
      7.竞技场体力：自动回复。
      8.普通体力：每天固定240+每日任务100+烤肉店72~120。
    
### 从产出层面讲
 
      1.简单难度关卡：金币、木、石、铁+少量装备升星材料。
      2.困难&地狱：金币、装备强化材料、少量装备升星材料。
      3.地牢：装备生星材料，金币。
      
## 三、普通抽奖的实现

      抽奖，实际上就是一个单纯的随机累加过程，每次讲随机到的东西数量+1，直到循环结束为止。
      分解与兑换装备也只是简单地计数加减乘除。
```markdown
If Sheets("Vow").Cells(17, 11) = "" Then
MsgBox ("Please enter the recharge amount!")
Else
End If

B = CLng(Sheets("Vow").Cells(17, 11) / (Sheets("Value").Cells(6, 3) / (Sheets("Value").Cells(3, 2) * Sheets("Value").Cells(3, 3))))
```

### Markdown

Markdown is a lightweight and easy-to-use syntax for styling your writing. It includes conventions for

```markdown
Syntax highlighted code block

# Header 1
## Header 2
### Header 3

- Bulleted
- List

1. Numbered
2. List

**Bold** and _Italic_ and `Code` text

[Link](url) and ![Image](src)
```

For more details see [GitHub Flavored Markdown](https://guides.github.com/features/mastering-markdown/).

### Jekyll Themes

Your Pages site will use the layout and styles from the Jekyll theme you have selected in your [repository settings](https://github.com/RogerDislikeMath/Roger--/settings). The name of this theme is saved in the Jekyll `_config.yml` configuration file.

### Support or Contact

Having trouble with Pages? Check out our [documentation](https://help.github.com/categories/github-pages-basics/) or [contact support](https://github.com/contact) and we’ll help you sort it out.
