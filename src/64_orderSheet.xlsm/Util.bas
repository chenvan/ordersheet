Attribute VB_Name = "Util"
Option Explicit

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

Function isInSheetRange(ByVal target As Variant, ByVal shName As String, ByVal rangeName As String) As Boolean
    Dim resultRng As range
    
    With Sheets(shName).range(rangeName)
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
        Sheets("回潮段").range("A3:A1171").ClearContents
        Sheets("回潮段").range("C3:K1171").ClearContents
        Sheets("回潮段").range("M3:N1171").ClearContents
        Sheets("回潮段").range("P3:P1171").ClearContents
        
        Sheets("加料段").range("A3:A1321").ClearContents
        Sheets("加料段").range("C3:D1321").ClearContents
        Sheets("加料段").range("G3:K1321").ClearContents
        Sheets("加料段").range("N3:P1321").ClearContents
        Sheets("加料段").range("R3:R1321").ClearContents
        
        Sheets("切烘加香段").range("A3:A1251").ClearContents
        Sheets("切烘加香段").range("C3:D1251").ClearContents
        Sheets("切烘加香段").range("G3:J1251").ClearContents
        Sheets("切烘加香段").range("L3:AC1251").ClearContents
        Sheets("切烘加香段").range("AE3:AE1251").ClearContents
        
        Sheets("HDT段").range("A3:A152").ClearContents
        Sheets("HDT段").range("C3:D152").ClearContents
        Sheets("HDT段").range("L3:L152").ClearContents
        
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
    Application.OnTime laterTime, "'showMsg """ & content & "'"
End Sub

Sub showMsg(ByVal content As String)
    Application.StatusBar = "##" & content & "   " & Left(Application.StatusBar, 80)
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
        '检查内容是否需要重置
        If tipSet.Exists("redirect") Then
            realContent = genNewContent(tobaccoName, tipSet("redirect"))
        ElseIf tipSet.Exists("hdt") Then
            '检查是否需要回掺HDT
            If getParam(tobaccoName, "HDT掺配比例") = "" Then GoTo Continue
            realContent = tipSet("hdt")
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
        pushVoiceTip realContent, realOffsetTime + delay, baseTime, tipSet("isForceBroadcast")
       
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
        Dim storeNamesInfeedLiquid As Variant
    
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
        
        pushVoiceTip tipSet("content"), realOffsetTime + delay, baseTime, tipSet("isForceBroadcast")
    
    Next tipSet
End Sub

Sub pushVoiceTip(ByVal content As String, ByVal tsOffset, ByVal baseTime As Variant, ByVal isForceBroadcast As Boolean)

    Dim triggerTime As Variant

    triggerTime = baseTime + TimeSerial(0, tsOffset, 2)

    If Time > triggerTime And isForceBroadcast Then
        speakLater Now, content
        showMsgLater Now, content
    ElseIf Time <= triggerTime Then
        speakLater triggerTime, content
        showMsgLater triggerTime, content
    Else
        showMsg "超时,不进行以下提醒: $" & content
    End If
End Sub

Sub runFirstBatchTip(ByVal sheetName As String)
    Dim found As range
    Dim firstTobaccoName As String
    'find today's first row
    Set found = Sheets(sheetName).range("A:A").Find(Date)
    
    If found Is Nothing Then
        Util.showMsg "没有找到今天的日期"
    Else
        firstTobaccoName = found.offset(0, 2).value
        '时间使用now的话会时间转换会溢出
        sheduleVoiceTips sheetName, firstTobaccoName, "第一批", Time, 0
    End If
End Sub

Public Function loadTips(ByVal fLayerP As String, ByVal sLayerP As String) As Variant
    Dim fullPath As String
    Dim allTips As Scripting.Dictionary
    
    fullPath = Sheets("设定").range("A:A").Find("语音文件路径").offset(0, 1).value
    
    Set allTips = loadJsonFile(fullPath)
    
    Set loadTips = allTips(fLayerP)(sLayerP)
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
    Dim rowIndex, columnIndex As Integer
    
    rowIndex = Sheets("设定").range("A2:A18").Find(tobaccoName).Row
    columnIndex = Sheets("设定").range("A1:Z1").Find(paramName).Column
    
    getParam = Sheets("设定").Cells(rowIndex, columnIndex)
End Function

Function genNewContent(ByVal tobaccoName As String, ByVal paramName As String) As String
    Dim param As String
    
    param = getParam(tobaccoName, paramName)
    
    genNewContent = paramName & param

End Function


Function adjustOffsetTime(ByVal tobaccoName As String, ByVal offsetTime As Integer) As Integer
    Dim mainTobaccoVolume As Integer
    
    mainTobaccoVolume = getParam(tobaccoName, "主叶丝秤流量")
    
    adjustOffsetTime = offsetTime * 6250 / mainTobaccoVolume
End Function



