Attribute VB_Name = "Util"
Option Explicit
Function ParseInputTime(ByVal abbrTime As Integer, ByVal dt As Variant) As Date
    
    Dim hour As Integer
    Dim min As Integer
    
    If dt = 0 Then
        MsgBox "请填写日期"
    ElseIf abbrTime > 2400 Then
        MsgBox "时间不能超过2400"
    Else
        hour = Int(abbrTime / 100)
        min = Int(abbrTime Mod 100)
    
        ParseInputTime = dt + TimeSerial(hour, min, 0)
    End If

End Function

Function IsInArray(ByVal beFound As Variant, ByRef arr As Variant) As Boolean
    Dim i As Integer
    
    For i = LBound(arr) To UBound(arr)
    
        If arr(i) = beFound Then
            IsInArray = True
            Exit Function
        End If
        
    Next i
    
    IsInArray = False
    
End Function

Function IsInCollection(ByVal beFound As Variant, ByRef coll As Variant) As Boolean
    Dim item As Variant
    
    For Each item In coll
        If item = beFound Then
            IsInCollection = True
            Exit Function
        End If
    Next
    
    IsInCollection = False
    
End Function

Function isInSheetRange(ByVal target As Variant, ByVal shName As String, ByVal rangeName As String) As Boolean
    Dim resultRng As Range
    
    With Sheets(shName).Range(rangeName)
        Set resultRng = .Find(What:=target)
        
        If resultRng Is Nothing Then
            isInSheetRange = False
        Else
            isInSheetRange = True
        End If
        
    End With
End Function

Function IsIntType(ByVal value As Variant) As Boolean
    'it seems Int function will translate empty string to 0
    If Int(value) = value Then
        IsIntType = True
    Else
        IsIntType = False
    End If
End Function


Function IsNum(ByVal value As Variant) As Boolean
    'IsNumeric function will let empty string be true
    IsNum = IsNumeric(value) And value <> ""
End Function

Sub clearContent()
    Dim answer As Integer
    Dim varResult As Variant
    
    answer = MsgBox("此操作将会把所有内容清空，是否已经把文件另存", vbYesNoCancel, "警告")
    
    If answer = 6 Then
        Sheets("回潮段").Range("A3:A1171").ClearContents
        Sheets("回潮段").Range("C3:J1171").ClearContents
        Sheets("回潮段").Range("L3:M1171").ClearContents
        Sheets("回潮段").Range("P3:P1171").ClearContents
        Sheets("回潮段").Range("P3:P1171").Font.Color = vbBlack
        
        Sheets("加料段").Range("A3:A1321").ClearContents
        Sheets("加料段").Range("C3:D1321").ClearContents
        Sheets("加料段").Range("F3:J1321").ClearContents
        Sheets("加料段").Range("M3:M1321").ClearContents
        Sheets("加料段").Range("O3:Q1321").ClearContents
        Sheets("加料段").Range("V3:V1321").ClearContents
        Sheets("加料段").Range("V3:V1321").Font.Color = vbBlack
        
        Sheets("切烘加香段").Range("A3:A1251").ClearContents
        Sheets("切烘加香段").Range("C3:D1251").ClearContents
        Sheets("切烘加香段").Range("F3:Y1251").ClearContents
        Sheets("切烘加香段").Range("AA3:AA1251").ClearContents
        Sheets("切烘加香段").Range("J3:J1251").Font.Color = vbBlack
        
        
        Sheets("HDT段").Range("A3:E152").ClearContents
       
    ElseIf answer = 7 Then
        varResult = Application.GetSaveAsFilename(filefilter:="Marco Enabled Workbook(*.xlsm), *xlsm")
        If varResult <> False Then
            ThisWorkbook.SaveAs (varResult)
        End If
    End If
End Sub

Sub speakLater(ByVal laterTime As Variant, ByVal content As String)
    Application.OnTime laterTime, "'speakAsync """ & content & "'"
End Sub

Sub speakAsync(ByVal content As String)
    Application.Speech.speak content & "。" & content, True
End Sub

Sub showMsgLater(ByVal laterTime As Variant, ByVal content As String)
    'https://stackoverflow.com/questions/31439866/multiple-variable-arguments-to-application-ontime
    Application.OnTime laterTime, "'showMsg """ & content & """,""" & laterTime & "'"
End Sub

Sub showMsg(ByVal content As String, Optional ByVal laterTime As Variant = "")
    
    If laterTime = "" Then
        laterTime = Now
    End If
    
    Application.StatusBar = Format(laterTime, "hh:mm") & " " & content & "   " & Left(Application.StatusBar, 80)
End Sub

Sub sheduleVoiceTips(ByVal sheetName As String, ByVal tobaccoName As String, ByVal producePhase As String, ByVal baseTime As Variant, ByVal delay As Integer)
    
    Dim tipArray As Variant
    Dim tipSet As Scripting.Dictionary
    Dim realContent As String
    Dim realOffsetTime As Variant
    
    '通过sheetName, producePhase加载所需的语音列表
    Set tipArray = loadTips(sheetName, producePhase)
    '遍历列表, 检查是否需要改变时间
    
    For Each tipSet In tipArray
        '检查 tipset 是否用在这牌号上
        If tipSet.Exists("filter") Then
            If Not IsInCollection(tobaccoName, tipSet("filter")) Then GoTo Continue
        End If
        
        '检查内容是否需要重置
        If tipSet.Exists("redirectList") Then
            realContent = genNewContent(tobaccoName, tipSet("redirectList"))
        Else
            realContent = tipSet("content")
        End If
        
        '检查时间是否需要重置
        If tipSet.Exists("aOffsetTime") Then
            realOffsetTime = adjustOffsetTime(tobaccoName, tipSet("aOffsetTime"))
        Else
            realOffsetTime = tipSet("sOffsetTime")
        End If
        
        'Debug.Print realContent
        '装载语音
        pushVoiceTip realContent, realOffsetTime + delay, baseTime, tipSet("deadLineOffset")
       
Continue:
    Next tipSet
    
End Sub

Sub sheduleVoiceTipsAboutStore(ByVal storePlace As String, ByVal tobaccoName As String, ByVal storeName As String, ByVal baseTime As Variant, ByVal delay As Integer)
    Dim tipArray As Variant
    Dim tipSet As Scripting.Dictionary
    Dim realOffsetTime As Variant
    Dim realStore As String
    
    If storePlace = "贮叶柜" Then
        '加载加料入柜tips
        Dim storesInFeedLiquid As Variant
    
        storesInFeedLiquid = Array("1", "2", "3", "4", "5", "6")
        realStore = "HDT"
        
        If IsInArray(storeName, storesInFeedLiquid) Then
            realStore = "叶柜"
        End If
        
    Else
        '加载切烘入柜tips
        Dim storesInAddEssence As Variant
        Dim storeIndex As Integer
        
        storesInAddEssence = Array("南A", "南B", "南C", "南D", "南E", "南F", "南G", "南H", "南J", "南K", "南L", "南M", "南N", "南P", "南Q", "南R", "南S", "南T", "木A", "木B", "木C", "北A", "北B", "北C", "北D", "北E", "北F", "北G", "北H", "外A", "外B")
        storeIndex = Application.Match(storeName, storesInAddEssence, False)
        
        realStore = "南AF"
        
        If storeIndex > 6 And storeIndex < 13 Then
            realStore = "南GM"
        ElseIf storeIndex > 12 And storeIndex <= 18 Then
            realStore = "南NT"
        ElseIf storeIndex > 18 And storeIndex <= 21 Then
            realStore = "木AC"
        ElseIf storeIndex > 21 And storeIndex <= 25 Then
            realStore = "北AD"
        ElseIf storeIndex > 25 And storeIndex <= 29 Then
            realStore = "北EH"
        ElseIf storeIndex > 29 And storeIndex <= 31 Then
            realStore = "外AB"
        End If
    End If
    
    Set tipArray = loadTips(storePlace, realStore)
    
    For Each tipSet In tipArray
        If tipSet.Exists("aOffsetTime") Then
            realOffsetTime = adjustOffsetTime(tobaccoName, tipSet("aOffsetTime"))
        Else
            realOffsetTime = tipSet("sOffsetTime")
        End If
        
        pushVoiceTip tipSet("content"), realOffsetTime + delay, baseTime, tipSet("deadLineOffset")
    
    Next tipSet
End Sub

Sub pushVoiceTip(ByVal content As String, ByVal tsOffset, ByVal baseTime As Variant, ByVal deadLineOffset As Integer)

    Dim triggerTime As Variant
    Dim deadLineTime As Variant
    
    triggerTime = baseTime + TimeSerial(0, tsOffset, 2)
    deadLineTime = triggerTime + TimeSerial(0, deadLineOffset, 0)

    If triggerTime < Now And Now < deadLineTime Then
        speakLater Now, content
        showMsgLater Now, content
    ElseIf Now <= triggerTime Then
        speakLater triggerTime, content
        showMsgLater triggerTime, content
    Else
        showMsg "超时,不进行提醒: $" & content
    End If
End Sub

'Sub runFirstBatchTip(ByVal sheetname As String)
'    '需要修改逻辑
'    Dim found As Range
'    Dim firstTobaccoName As String
'    'find today's first row
'    Set found = Sheets(sheetname + "段").Range("A:A").Find(Date)
'
'    If found Is Nothing Then
'        Util.showMsg "没有找到今天的日期"
'    Else
'        firstTobaccoName = found.Offset(0, 2).value
'        sheduleVoiceTips sheetname, firstTobaccoName, "第一批", Now, 0
'    End If
'End Sub

Sub checkBeforeWork(ByVal sheetName As String)
    Dim found As Range
    Dim firstTobaccoName As String
    'find today's first row
    Set found = Sheets(sheetName + "段").Range("A:A").Find(Date)

    If found Is Nothing Then
        Util.showMsg "没有找到今天的日期"
    Else
        firstTobaccoName = found.Offset(0, 2).value
        sheduleVoiceTips sheetName, "", "第一批", Now, 0
        sheduleVoiceTips sheetName, firstTobaccoName, "开始前", Now, 0
    End If
End Sub

Public Function loadTips(ByVal fLayerP As String, ByVal sLayerP As String) As Variant
    On Error GoTo EH
    
    Dim fullPath As String
    Dim allTips As Scripting.Dictionary
    
    fullPath = Sheets("设定").Range("A:A").Find("语音文件路径").Offset(0, 1).value
    
    Set allTips = loadJsonFile(fullPath)
    Set loadTips = allTips(fLayerP)(sLayerP)
    
    Exit Function
EH:
    MsgBox Err.Description & vbCrLf & "在JSON文件中无法找到: " & fLayerP & " -> " & sLayerP & " 属性"
End Function

Function loadJsonFile(ByVal path As String) As Variant
    Dim fso As New FileSystemObject
    Dim JsonTS As TextStream
    Dim JsonText As String
    
    Set JsonTS = fso.OpenTextFile(path, ForReading)
    JsonText = JsonTS.ReadAll
    JsonTS.Close
    
    Set loadJsonFile = JsonConverter.ParseJson(JsonText)
End Function

Function getParam(ByVal tobaccoName As String, ByVal paramName As String) As Variant
    On Error GoTo EH
    
    Dim rowIndex, columnIndex As Integer
 
    rowIndex = Sheets("设定").Range("A2:A25").Find(tobaccoName, lookat:=xlWhole).Row
    columnIndex = Sheets("设定").Range("A1:Z1").Find(paramName, lookat:=xlWhole).Column
    
    getParam = Sheets("设定").Cells(rowIndex, columnIndex)
    
    Exit Function
EH:
    MsgBox Err.Description & vbCrLf & "在设定表中无法找到: " & tobaccoName & " -> " & paramName & " 参数"
End Function


Function getColumnIndex(ByVal sheetName As String, ByVal targetCol As String) As Integer

    getColumnIndex = Sheets(sheetName).Range("A2:AZ2").Find(targetCol, lookat:=xlWhole).Column

End Function

Function genNewContent(ByVal tobaccoName As String, ByRef paramCollection As Variant) As String
    Dim newContent, param, paramContent As String
    newContent = tobaccoName
    
    For Each param In paramCollection
        newContent = newContent & "," & param & getParam(tobaccoName, param)
    Next
    
    genNewContent = newContent

End Function


Function adjustOffsetTime(ByVal tobaccoName As String, ByVal offsetTime As Integer) As Integer
    Dim mainTobaccoVolume As Integer
    
    mainTobaccoVolume = getParam(tobaccoName, "主叶丝秤流量")
    '偏移时间是以经典主叶丝秤流量为基准的, 其他牌号通过自己的流量换算出自己的偏移时间
    '掺配和出现不同的偏移时间是因为主叶丝秤流量不同
    adjustOffsetTime = offsetTime + (9.375 - 0.0015 * mainTobaccoVolume)
End Function

Function countSequentialTobaccoNum(ByVal tobaccoName As String, ByVal sheetName As String, ByVal dateColIndex As Integer, ByVal tobaccoColIndex As Integer, ByVal lastRowIndex As Integer) As Integer
    Dim prevTobaccoName As String
    Dim currentDate, prevDate As Variant
    Dim count As Integer
    
    count = 0
    
    prevTobaccoName = Sheets(sheetName).Cells(lastRowIndex, tobaccoColIndex).value
    prevDate = Sheets(sheetName).Cells(lastRowIndex, dateColIndex).value
    currentDate = Sheets(sheetName).Cells(lastRowIndex + 1, dateColIndex).value
    
    While tobaccoName = prevTobaccoName And currentDate = prevDate
        count = count + 1
        prevTobaccoName = Sheets(sheetName).Cells(lastRowIndex - count, tobaccoColIndex).value
        prevDate = Sheets(sheetName).Cells(lastRowIndex - count, dateColIndex).value
    Wend
    
    countSequentialTobaccoNum = count
End Function

Function getAddEssenceSweepDelay(ByVal tobaccoName As String, ByVal dateColIndex As Integer, ByVal tobaccoColIndex As Integer, ByVal serialNumColIndex As Integer, ByVal lastRowIndex As Integer) As Integer
    Dim count, delay, serialNum As Integer

    count = countSequentialTobaccoNum(tobaccoName, "切烘加香段", dateColIndex, tobaccoColIndex, lastRowIndex)
    delay = 8
    
    If count Mod 3 = 0 Then
        '三批同牌号清洗一次
        delay = 8
    ElseIf count > 0 Then
        serialNum = Sheets("切烘加香段").Cells(lastRowIndex, serialNumColIndex).value
        '第4批和第8批加香机需要清扫, 延时4分钟上烟
        If serialNum = 4 Or serialNum = 8 Then
            delay = 4
        Else
            delay = 0
        End If
    End If
'    Debug.Print count
'    Debug.Print delay
    getAddEssenceSweepDelay = delay
End Function

Function getFeedLiquidSweepDelay(ByVal tobaccoName As String, ByVal dateColIndex As Integer, ByVal tobaccoColIndex As Integer, ByVal lastRowIndex As Integer) As Integer
    Dim count, delay As Integer
    
    count = countSequentialTobaccoNum(tobaccoName, "加料段", dateColIndex, tobaccoColIndex, lastRowIndex)
    delay = 10
    
    If count Mod 4 = 0 Then
        '四批同牌号清洗一次
        delay = 10
    ElseIf count > 0 Then
        delay = 0
    End If
    
    getFeedLiquidSweepDelay = delay
    
End Function

Function guessFinishTime(sheetName As String) As String

    Dim endTimeColIndex, tobaccoNameColIndex As Integer
    Dim tobaccoNameLastRow, endTimeLastRow As Integer
    Dim switchBaseTime, switchAddTime, tobaccoFinishTime As Integer
    Dim index As Integer
    Dim beginTime, endTime As Date
    Dim output, tobaccoName As String

    endTimeColIndex = Util.getColumnIndex(sheetName, "结束时间")
    '牌号的 colIndex 基本不会变, 都是第三列
    tobaccoNameColIndex = 3
    output = ""
    switchBaseTime = Sheets("设定").Range("A:A").Find(sheetName + "转烟时间").Offset(0, 1).value

    tobaccoNameLastRow = Cells(Rows.count, tobaccoNameColIndex).End(xlUp).Row + 1
    endTimeLastRow = Cells(Rows.count, endTimeColIndex).End(xlUp).Row + 1

    tobaccoName = Cells(endTimeLastRow, tobaccoNameColIndex).value

    If Cells(endTimeLastRow, endTimeColIndex - 1).value = "" Then
        'beginTime is empty
        switchAddTime = 0
        If sheetName = "加料段" Then
            switchAddTime = Util.getFeedLiquidSweepDelay(tobaccoName, 1, 3, endTimeLastRow - 1)
        ElseIf sheetName = "切烘加香段" Then
            switchAddTime = Util.getAddEssenceSweepDelay(tobaccoName, 1, 3, 2, endTimeLastRow - 1)
        End If
        beginTime = DateAdd("n", switchBaseTime + switchAddTime, Cells(endTimeLastRow - 1, endTimeColIndex).value)
    Else
        beginTime = Cells(endTimeLastRow, endTimeColIndex - 1).value
    End If


    For index = endTimeLastRow To tobaccoNameLastRow - 1

        tobaccoFinishTime = Util.getParam(tobaccoName, sheetName + "生产时长")

        endTime = DateAdd("n", tobaccoFinishTime, beginTime)

        output = output + tobaccoName + " : " + Format(beginTime, "hh:mm") + " - " + Format(endTime, "hh:mm") + vbCrLf

        tobaccoName = Cells(index + 1, tobaccoNameColIndex).value

        switchAddTime = 0

        If sheetName = "加料段" Then
            switchAddTime = Util.getFeedLiquidSweepDelay(tobaccoName, 1, 3, index)
        ElseIf sheetName = "切烘加香段" Then
            switchAddTime = Util.getAddEssenceSweepDelay(tobaccoName, 1, 3, 2, index)
        End If

        beginTime = DateAdd("n", switchBaseTime + switchAddTime, endTime)

'        Debug.Print switchAddTime
'        Debug.Print beginTime
    Next

    guessFinishTime = output

End Function



