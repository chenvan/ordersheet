Attribute VB_Name = "Util"
Private Type tipInfo
    tipContent As String
    tipTriggerTime As Variant
    tipTimeDiffInMin As Integer
End Type

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


Sub shedule(ByVal target As String, ByVal tAnchor As Variant, ByVal cOffset As Integer, ByVal delay As Integer, Optional ByVal isTriggerNow As Boolean = False)
    '语音提示安排函数
    '@taget: 烟牌号
    '@tAnchor: 时间锚点
    '@cOffset: 所属生产段与A列的列间隔
    '@delay: 推迟时间
    '@isTriggerNow: 立即提醒
    
    Dim found As range
    Dim col As String
    Dim index, lastRow As Integer
    Dim info As tipInfo
    
    Set found = Sheets("语音提示").range("A:A").Find(target, , , xlWhole)
    col = found.offset(0, cOffset).value
    
    With Sheets("语音提示")
        '找到语音提醒的最后一行
        'use .End(xlDown).End(xlDown).End(xlUp) instead of only use .End(xlDown)
        'it can avoid return too many rowscount if the tips is empty
        lastRow = .range(col & "1", .range(col & "1").End(xlDown).End(xlDown).End(xlUp)).Rows.Count
        'Debug.Print lastRow
        'Debug.Print tAnchor
        
        '语音提醒是从第二行开始
        'Debug.Print "last row: " & lastRow
        
        For index = 2 To lastRow

            info = getTipInfo(col, index, tAnchor, delay, isTriggerNow)
            runTipInfo info, -11
            
        Next index
    
    End With
       
End Sub

Sub speakPrecaution(ByVal target As String, ByVal tAnchor As Variant, ByVal colName As String, ByVal delay As Integer, Optional ByVal isTriggerNow As Boolean = False)
    Dim found As range
    Dim preCautionTip As tipInfo

    Set found = Sheets("语音提示").range("A:A").Find(target, , , xlWhole)
    preCautionTip = getTipInfo(colName, found.Row, tAnchor, delay, isTriggerNow)

    runTipInfo preCautionTip, -11

End Sub

Private Function getTipInfo(ByVal colName As String, ByVal rowIndex As Integer, ByVal tAnchor As Variant, ByVal delay As Integer, ByVal isTriggerNow As Boolean) As tipInfo
    Dim info As tipInfo
    Dim tsOffset As Integer
    Dim triggerTime As Variant
    Dim content As String
    
    '语音提醒时间与锚点的时间间隔

    tsOffset = Sheets("语音提示").range(colName & rowIndex).value + delay
    content = Sheets("语音提示").range(colName & rowIndex).offset(0, 1).value
    
    If isTriggerNow Then
        triggerTime = Time + TimeSerial(0, 0, 2)
    Else
        triggerTime = tAnchor + TimeSerial(0, tsOffset, 0)
    End If

    With info
        .tipContent = content
        .tipTriggerTime = triggerTime
        .tipTimeDiffInMin = (triggerTime - Time) * 1440
    End With
    
    getTipInfo = info
End Function

Sub runTipInfo(ByRef info As tipInfo, ByVal latestTime As Integer)
    If info.tipTimeDiffInMin > latestTime Then
       speakLater info.tipTriggerTime, info.tipContent
       showMsgLater info.tipTriggerTime, info.tipContent
    Else
        showMsg "超时大于" & Abs(latestTime) - 1 & "分钟, 不会进行语音提醒, $" & info.tipContent
    End If
End Sub

Sub runFirstBatchWarning(ByVal sheetName, ByVal colNameForBeforeBegin, ByVal colNameForPrecaution)
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
        Util.shedule firstTobaccoName, Time, colNameForBeforeBegin, 0, True
        Util.speakPrecaution firstTobaccoName, Time, colNameForPrecaution, 0, True
    End If
End Sub









