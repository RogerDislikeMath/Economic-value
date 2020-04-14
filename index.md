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

### 抽奖，实际上就是一个单纯的随机累加过程，每次讲随机到的东西数量+1，直到循环结束为止。分解与兑换装备也只是简单地计数加减乘除。
```markdown
If Sheets("Vow").Cells(17, 11) = "" Then
MsgBox ("Please enter the recharge amount!")
Else
End If

B = CLng(Sheets("Vow").Cells(17, 11) / (Sheets("Value").Cells(6, 3) / (Sheets("Value").Cells(3, 2) * Sheets("Value").Cells(3, 3))))
```

### 每次中奖次数+1，比较容易理解


```markdown
For N = 1 To B
M = Rnd()
If M <= Sheets("Vow").Cells(8, 7) Then
 Sheets("Vow").Cells(14, 16) = Sheets("Vow").Cells(14, 16) + Sheets("Vow").Cells(7, 7)
ElseIf M <= (Sheets("Vow").Cells(8, 7) + Sheets("Vow").Cells(8, 8)) Then
 Sheets("Vow").Cells(15, 16) = Sheets("Vow").Cells(15, 16) + Sheets("Vow").Cells(7, 8)
ElseIf M <= (Sheets("Vow").Cells(8, 7) + Sheets("Vow").Cells(8, 8) + Sheets("Vow").Cells(8, 9)) Then
 Sheets("Vow").Cells(16, 16) = Sheets("Vow").Cells(16, 16) + Sheets("Vow").Cells(7, 9)
ElseIf M <= (Sheets("Vow").Cells(8, 7) + Sheets("Vow").Cells(8, 8) + Sheets("Vow").Cells(8, 9) + Sheets("Vow").Cells(8, 10)) Then
 Sheets("Vow").Cells(17, 16) = Sheets("Vow").Cells(17, 16) + Sheets("Vow").Cells(7, 10)
ElseIf M <= (Sheets("Vow").Cells(8, 7) + Sheets("Vow").Cells(8, 8) + Sheets("Vow").Cells(8, 9) + Sheets("Vow").Cells(8, 10) + Sheets("Vow").Cells(8, 11)) Then
 Sheets("Vow").Cells(18, 16) = Sheets("Vow").Cells(18, 16) + Sheets("Vow").Cells(7, 11)
ElseIf M <= (Sheets("Vow").Cells(8, 7) + Sheets("Vow").Cells(8, 8) + Sheets("Vow").Cells(8, 9) + Sheets("Vow").Cells(8, 10) + Sheets("Vow").Cells(8, 11) + Sheets("Vow").Cells(8, 12)) Then
 Sheets("Vow").Cells(19, 16) = Sheets("Vow").Cells(19, 16) + Sheets("Vow").Cells(7, 12)
ElseIf M <= (Sheets("Vow").Cells(8, 7) + Sheets("Vow").Cells(8, 8) + Sheets("Vow").Cells(8, 9) + Sheets("Vow").Cells(8, 10) + Sheets("Vow").Cells(8, 11) + Sheets("Vow").Cells(8, 12) + Sheets("Vow").Cells(8, 13)) Then
 Sheets("Vow").Cells(20, 16) = Sheets("Vow").Cells(20, 16) + Sheets("Vow").Cells(7, 13)
ElseIf M <= (Sheets("Vow").Cells(8, 7) + Sheets("Vow").Cells(8, 8) + Sheets("Vow").Cells(8, 9) + Sheets("Vow").Cells(8, 10) + Sheets("Vow").Cells(8, 11) + Sheets("Vow").Cells(8, 12) + Sheets("Vow").Cells(8, 13) + Sheets("Vow").Cells(8, 14)) Then
 Sheets("Vow").Cells(21, 16) = Sheets("Vow").Cells(21, 16) + Sheets("Vow").Cells(7, 14)
ElseIf M <= (Sheets("Vow").Cells(8, 7) + Sheets("Vow").Cells(8, 8) + Sheets("Vow").Cells(8, 9) + Sheets("Vow").Cells(8, 10) + Sheets("Vow").Cells(8, 11) + Sheets("Vow").Cells(8, 12) + Sheets("Vow").Cells(8, 13) + Sheets("Vow").Cells(8, 14) + Sheets("Vow").Cells(8, 15)) Then
 Sheets("Vow").Cells(22, 16) = Sheets("Vow").Cells(22, 16) + Sheets("Vow").Cells(7, 15)
Else
 Sheets("Vow").Cells(23, 16) = Sheets("Vow").Cells(23, 16) + Sheets("Vow").Cells(7, 16)
End If
```

### 将结果同步到后面分解部分

```markdown
Sheets("Vow").Cells(27, 16) = Sheets("Vow").Cells(14, 16)
Sheets("Vow").Cells(28, 16) = Sheets("Vow").Cells(15, 16)
Sheets("Vow").Cells(29, 16) = Sheets("Vow").Cells(16, 16)
Sheets("Vow").Cells(30, 16) = Sheets("Vow").Cells(17, 16)
Sheets("Vow").Cells(31, 16) = Sheets("Vow").Cells(18, 16)
Sheets("Vow").Cells(32, 16) = Sheets("Vow").Cells(19, 16)
Sheets("Vow").Cells(33, 16) = Sheets("Vow").Cells(20, 16)
Sheets("Vow").Cells(34, 16) = Sheets("Vow").Cells(21, 16)
Sheets("Vow").Cells(35, 16) = Sheets("Vow").Cells(22, 16)
Sheets("Vow").Cells(36, 16) = Sheets("Vow").Cells(23, 16)
```

## 三、关卡掉落部分实现

### VBA不具备python字典功能，因此使用IF-Else进行小函数构造。

```markdown
Function 难度(X)
    If X = "普通" Then
        难度 = 1
    ElseIf X = "困难" Then
        难度 = 2
    ElseIf X = "地狱" Then
        难度 = 3
    Else
    End If
End Function
Function 岛屿(X)
    If X = "岛屿1" Then
        岛屿 = 1
    ElseIf X = "岛屿2" Then
        岛屿 = 2
    ElseIf X = "岛屿3" Then
        岛屿 = 3
    ElseIf X = "岛屿4" Then
        岛屿 = 4
    ElseIf X = "岛屿5" Then
        岛屿 = 5
    ElseIf X = "岛屿6" Then
        岛屿 = 6
    ElseIf X = "岛屿7" Then
        岛屿 = 7
    ElseIf X = "岛屿8" Then
        岛屿 = 8
    ElseIf X = "岛屿9" Then
        岛屿 = 9
    Else
    End If
End Function

```

### 这里决定了要循环到多少行，具体数量可根据公式在Excel中自动计算得出。
### 整体循环顺序为大循环为各关卡(自上而下)
### 小循环为每关卡的掉落

```markdown
行 = 2
For Checkpoint = 3 To Sheets("DropTable").Cells(2, 18)
Num = 1

```
### 6-9列掉落设置为金币、木材、石头、铁，因此该部分为概率掉落

```markdown
For DropTable = 6 To 9
If Sheets("ActualDrop").Cells(Checkpoint, DropTable) = "" Then

Else
Sheets("DropTable").Cells(行, 1) = 难度(Sheets("ActualDrop").Cells(Checkpoint, 1)) * 1000000 + 岛屿(Sheets("ActualDrop").Cells(Checkpoint, 2)) * 100000 + Sheets("ActualDrop").Cells(Checkpoint, 3) * 100 + Num
Sheets("DropTable").Cells(行, 3) = 1
Sheets("DropTable").Cells(行, 4) = 3
Sheets("DropTable").Cells(行, 7) = 1
Sheets("DropTable").Cells(行, 8) = Sheets("ActualDrop").Cells(1, DropTable)
Sheets("DropTable").Cells(行, 9) = 99999999
Sheets("DropTable").Cells(行, 10) = Sheets("ActualDrop").Cells(Checkpoint, DropTable)
Sheets("DropTable").Cells(行, 11) = 1000
行 = 行 + 1
Num = Num + 1
End If
Next DropTable

```

### 11-13列掉落设置为1~3级强化石，且概率读取第10列，因此代码如下

```markdown
For DropTable = 11 To 13
If Sheets("ActualDrop").Cells(Checkpoint, DropTable) = "" Then

Else
Sheets("DropTable").Cells(行, 1) = 难度(Sheets("ActualDrop").Cells(Checkpoint, 1)) * 1000000 + 岛屿(Sheets("ActualDrop").Cells(Checkpoint, 2)) * 100000 + Sheets("ActualDrop").Cells(Checkpoint, 3) * 100 + Num
Sheets("DropTable").Cells(行, 3) = 1
Sheets("DropTable").Cells(行, 4) = 2
Sheets("DropTable").Cells(行, 7) = 1
Sheets("DropTable").Cells(行, 8) = Sheets("ActualDrop").Cells(1, DropTable)
Sheets("DropTable").Cells(行, 9) = 99999999
Sheets("DropTable").Cells(行, 10) = Sheets("ActualDrop").Cells(Checkpoint, DropTable)
Sheets("DropTable").Cells(行, 12) = Sheets("ActualDrop").Cells(Checkpoint, 10) * 100
行 = 行 + 1
Num = Num + 1
End If
Next DropTable

```

### 后面均同理

```markdown
For DropTable = 15 To 20
If Sheets("ActualDrop").Cells(Checkpoint, DropTable) = "" Then

Else
Sheets("DropTable").Cells(行, 1) = 难度(Sheets("ActualDrop").Cells(Checkpoint, 1)) * 1000000 + 岛屿(Sheets("ActualDrop").Cells(Checkpoint, 2)) * 100000 + Sheets("ActualDrop").Cells(Checkpoint, 3) * 100 + Num
Sheets("DropTable").Cells(行, 3) = 1
Sheets("DropTable").Cells(行, 4) = 2
Sheets("DropTable").Cells(行, 7) = 1
Sheets("DropTable").Cells(行, 8) = Sheets("ActualDrop").Cells(1, DropTable)
Sheets("DropTable").Cells(行, 9) = 99999999
Sheets("DropTable").Cells(行, 10) = Sheets("ActualDrop").Cells(Checkpoint, DropTable)
Sheets("DropTable").Cells(行, 12) = Sheets("ActualDrop").Cells(Checkpoint, 14) * 100
行 = 行 + 1
Num = Num + 1
End If
Next DropTable

For DropTable = 22 To 27
If Sheets("ActualDrop").Cells(Checkpoint, DropTable) = "" Then

Else
Sheets("DropTable").Cells(行, 1) = 难度(Sheets("ActualDrop").Cells(Checkpoint, 1)) * 1000000 + 岛屿(Sheets("ActualDrop").Cells(Checkpoint, 2)) * 100000 + Sheets("ActualDrop").Cells(Checkpoint, 3) * 100 + Num
Sheets("DropTable").Cells(行, 3) = 1
Sheets("DropTable").Cells(行, 4) = 2
Sheets("DropTable").Cells(行, 7) = 1
Sheets("DropTable").Cells(行, 8) = Sheets("ActualDrop").Cells(1, DropTable)
Sheets("DropTable").Cells(行, 9) = 99999999
Sheets("DropTable").Cells(行, 10) = Sheets("ActualDrop").Cells(Checkpoint, DropTable)
Sheets("DropTable").Cells(行, 12) = Sheets("ActualDrop").Cells(Checkpoint, 21) * 100
行 = 行 + 1
Num = Num + 1
End If
Next DropTable

For DropTable = 29 To 34
If Sheets("ActualDrop").Cells(Checkpoint, DropTable) = "" Then

Else
Sheets("DropTable").Cells(行, 1) = 难度(Sheets("ActualDrop").Cells(Checkpoint, 1)) * 1000000 + 岛屿(Sheets("ActualDrop").Cells(Checkpoint, 2)) * 100000 + Sheets("ActualDrop").Cells(Checkpoint, 3) * 100 + Num
Sheets("DropTable").Cells(行, 3) = 1
Sheets("DropTable").Cells(行, 4) = 2
Sheets("DropTable").Cells(行, 7) = 1
Sheets("DropTable").Cells(行, 8) = Sheets("ActualDrop").Cells(1, DropTable)
Sheets("DropTable").Cells(行, 9) = 99999999
Sheets("DropTable").Cells(行, 10) = Sheets("ActualDrop").Cells(Checkpoint, DropTable)
Sheets("DropTable").Cells(行, 12) = Sheets("ActualDrop").Cells(Checkpoint, 28) * 100
行 = 行 + 1
Num = Num + 1
End If
Next DropTable

For DropTable = 36 To 41
If Sheets("ActualDrop").Cells(Checkpoint, DropTable) = "" Then

Else
Sheets("DropTable").Cells(行, 1) = 难度(Sheets("ActualDrop").Cells(Checkpoint, 1)) * 1000000 + 岛屿(Sheets("ActualDrop").Cells(Checkpoint, 2)) * 100000 + Sheets("ActualDrop").Cells(Checkpoint, 3) * 100 + Num
Sheets("DropTable").Cells(行, 3) = 1
Sheets("DropTable").Cells(行, 4) = 2
Sheets("DropTable").Cells(行, 7) = 1
Sheets("DropTable").Cells(行, 8) = Sheets("ActualDrop").Cells(1, DropTable)
Sheets("DropTable").Cells(行, 9) = 99999999
Sheets("DropTable").Cells(行, 10) = Sheets("ActualDrop").Cells(Checkpoint, DropTable)
Sheets("DropTable").Cells(行, 12) = Sheets("ActualDrop").Cells(Checkpoint, 35) * 100
行 = 行 + 1
Num = Num + 1
End If
Next DropTable

```

### 掉落包的部分，难点在于需要额外加通用掉落(装备兑换券)
```markdown

Sheets("DropTable").Cells(行, 1) = 难度(Sheets("ActualDrop").Cells(Checkpoint, 1)) * 10000 + 岛屿(Sheets("ActualDrop").Cells(Checkpoint, 2)) * 1000 + Sheets("ActualDrop").Cells(Checkpoint, 3)
Sheets("DropTable").Cells(行, 2) = Sheets("DropTable").Cells(行 - Num + 1, 8)
Sheets("DropTable").Cells(行, 3) = 2
Sheets("DropTable").Cells(行, 4) = 4
Sheets("DropTable").Cells(行, 7) = 1

For DropTable = 1 To Num - 1

If DropTable >= 2 Then
Sheets("DropTable").Cells(行 - DropTable + 1, 2) = Sheets("DropTable").Cells(行 - DropTable + 1, 2) & "," & Sheets("DropTable").Cells(行 - Num + 1, 8)
Else
End If
Sheets("DropTable").Cells(行, 8) = Sheets("DropTable").Cells(行 - Num + 1, 1)
Sheets("DropTable").Cells(行, 12) = 100
行 = 行 + 1

If DropTable = (Num - 1) Then
Sheets("DropTable").Cells(行, 8) = 1000000 + 难度(Sheets("ActualDrop").Cells(Checkpoint, 1))
Sheets("DropTable").Cells(行, 12) = 1
行 = 行 + 1
Else
End If
Next DropTable
```
