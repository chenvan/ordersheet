Attribute VB_Name = "Util"
'Option Explicit


Function IsInArray(beFound As Variant, arr As Variant) As Boolean
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
    '如果安排的时间已经过了,应该立即进行语音提醒
    'Debug.Print Time
    'Debug.Print laterTime
    If Time >= laterTime Then
        'Debug.Print "超时"
        content = "超时," & content
        Application.OnTime Now, "'speakAsync """ & content & "'"
    Else
        Application.OnTime laterTime, "'speakAsync """ & content & "'"
    End If
End Sub

Sub speakAsync(ByVal content As String)
    Application.Speech.speak content & "。" & content, True
End Sub

Sub showMsgLater(ByVal laterTime As Variant, ByVal content As String)
    If Time >= laterTime Then
        content = "超时," & content
        Application.OnTime Now, "'showMsg """ & content & "'"
    Else
        Application.OnTime laterTime, "'showMsg """ & content & "'"
    End If
End Sub

Sub showMsg(ByVal content As String)
    Application.StatusBar = "##" & content & "   " & Left(Application.StatusBar, 80)
End Sub

Sub shedule(ByVal sheetName As String, ByVal tobaccoName As String, ByVal producePhase As String, ByVal baseTime As Variant, ByVal delay As Integer)
    Dim tipPair As Object
    Dim tobaccoTips, defaultTips As Object
    
    Set defaultTips = Util.loadDefaultTips
    'Util.showMsg "载入 default 提示文件"
    
    For Each tipPair In defaultTips(sheetName)(producePhase)
        'Debug.Print tipPair("延时") & ", " & tipPair("内容")
        loadTip tipPair("内容"), tipPair("延时") + delay, baseTime
    Next tipPair
    
    'load tobacco tips
    Set tobaccoTips = Util.loadTobaccoTips(tobaccoName)
    'Util.showMsg "载入 " & tobaccoName & " 提示文件"

    For Each tipPair In tobaccoTips(sheetName)(producePhase)
        'Debug.Print tipPair("延时") & ", " & tipPair("内容")
        loadTip tipPair("内容"), tipPair("延时") + delay, baseTime
    Next tipPair
    
End Sub

Sub sheduleDestTips(ByVal sheetName As String, ByVal cabinetName As String, ByVal baseTime As Variant)
    Dim tipPair As Object
    Dim destTips As Variant
    
    Set destTips = loadCabinetTips(cabinetName, sheetName)
    'Util.showMsg "载入 cabinet 提示文件"
    
    For Each tipPair In destTips
        'Debug.Print tipPair("延时") & ", " & tipPair("内容")
        loadTip tipPair("内容"), tipPair("延时"), baseTime
    Next tipPair
End Sub


Sub loadTip(ByVal content As String, ByVal tsOffset, ByVal baseTime As Variant)

    Dim triggerTime As Variant
    Dim timeDiffInMin As Integer
    Dim latestTime As Integer
    
    triggerTime = baseTime + TimeSerial(0, tsOffset, 2)
    timeDiffInMin = (triggerTime - Time) * 1440
    latestTime = -11
    
    If timeDiffInMin > latestTime Then
        speakLater triggerTime, content
        showMsgLater triggerTime, content
    Else
        showMsg "超时大于" & Abs(latestTime) - 1 & "分钟, 不会进行语音提醒, $" & content
    End If
End Sub

Sub runFirstBatchTip(ByVal sheetName As String)
    Dim found As range
    Dim firstTobaccoName As String
    'find today's first row
    Set found = Sheets(sheetName).range("A:A").Find(Date, , , xlWhole)
    If found Is Nothing Then
        Util.showMsg "没有找到今天的日期"
    Else
        'get the tobacco name, run util.speakPrecation & 开始前提醒
        firstTobaccoName = found.offset(0, 2).value
        'Debug.Print firstTobaccoName
        '时间使用now的话会时间转换会溢出
'        Debug.Print sheetName
'        Debug.Print firstTobaccoName
        Util.shedule sheetName, firstTobaccoName, "第一批", Time, 0
    End If
End Sub

Public Function loadDefaultTips() As Object
    Dim path As String

    path = getBasePathOfTips() & "\defaultTips.json"

    Set loadDefaultTips = loadJsonFile(path)

End Function

Function loadTobaccoTips(ByVal tobaccoName As String) As Object
    Dim path As String
    Dim allTobaccoTips As Object
    
    path = getBasePathOfTips() & "\tobaccoTips.json"
    
    Set allTobaccoTips = loadJsonFile(path)
    
    Set loadTobaccoTips = allTobaccoTips(tobaccoName)
End Function

Function loadCabinetTips(ByVal cabinetName As String, ByVal sheetName As String) As Variant
    Dim path As String
    Dim cabinetTips As Object
    Dim mark As String
    
    path = getBasePathOfTips() & "\cabinetTips.json"
    
    Set cabinetTips = loadJsonFile(path)
    mark = cabinetTips(sheetName)(cabinetName)
    
    '返回的是array
    Set loadCabinetTips = cabinetTips(sheetName)(mark)
End Function

Function loadJsonFile(ByVal path As String) As Object
    Dim fso As New FileSystemObject
    Dim JsonTS As TextStream
    Dim JsonText As String
    
    Set JsonTS = fso.OpenTextFile(path, ForReading)
    JsonText = JsonTS.ReadAll
    JsonTS.Close
    
    Set loadJsonFile = JsonConverter.ParseJson(JsonText)
End Function

Function findInColumn(sheetName As String, rangeName As String, target As String) As range
    Set findInColumn = Sheets(sheetName).range(rangeName).Find(target, , , xlWhole)
End Function

Function getBasePathOfTips() As String
    Dim found As range
    Set found = findInColumn("设定", "A:A", "语音文件路径")
    
    getBasePathOfTips = found.offset(0, 1).value
End Function










