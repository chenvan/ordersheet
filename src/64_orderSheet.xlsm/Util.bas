Attribute VB_Name = "Util"
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
        Sheets("回潮段").Range("C3:K1171").ClearContents
        Sheets("回潮段").Range("M3:N1171").ClearContents
        Sheets("回潮段").Range("P3:P1171").ClearContents
        
        Sheets("加料段").Range("A3:A1321").ClearContents
        Sheets("加料段").Range("C3:D1321").ClearContents
        Sheets("加料段").Range("G3:K1321").ClearContents
        Sheets("加料段").Range("N3:P1321").ClearContents
        Sheets("加料段").Range("R3:R1321").ClearContents
        
        Sheets("切烘加香段").Range("A3:A1251").ClearContents
        Sheets("切烘加香段").Range("C3:D1251").ClearContents
        Sheets("切烘加香段").Range("G3:J1251").ClearContents
        Sheets("切烘加香段").Range("L3:AC1251").ClearContents
        Sheets("切烘加香段").Range("AE3:AE1251").ClearContents
        
        Sheets("HDT段").Range("A3:A152").ClearContents
        Sheets("HDT段").Range("C3:D152").ClearContents
        Sheets("HDT段").Range("L3:L152").ClearContents
        
    ElseIf answer = 7 Then
        varResult = Application.GetSaveAsFilename(filefilter:="Marco Enabled Workbook(*.xlsm), *xlsm")
        If varResult <> False Then
            ThisWorkbook.SaveAs (varResult)
        End If
    End If
End Sub

Sub speakLater(ByVal laterTime As Variant, ByVal content As String)
    '如果安排的时间已经过了,应该立即进行语音提醒
    If Now() >= laterTime Then
        content = "已超时," & content
        Application.OnTime Now(), "'speakAsync """ & content & "'"
    Else
        Application.OnTime laterTime, "'speakAsync """ & content & "'"
    End If
End Sub

Sub speakAsync(ByVal content As String)
    Application.Speech.speak content & "。" & content, True
End Sub

Sub showMsgLater(ByVal laterTime As Variant, ByVal content As String)
    If Now() >= laterTime Then
        content = "已超时," & content
        Application.OnTime Now(), "'showMsg """ & content & "'"
    Else
        Application.OnTime laterTime, "'showMsg """ & content & "'"
    End If
End Sub

Sub showMsg(ByVal content As String)
    Application.StatusBar = "##" & content & "   " & Left(Application.StatusBar, 80)
End Sub


Sub shedule(ByVal target As String, ByVal tAnchor As Variant, ByVal cOffset As Integer, ByVal delay As Integer)
    '语音提示安排函数
    '@taget: 烟牌号
    '@tAnchor: 时间锚点
    '@cOffset: 所属生产段与A列的列间隔
    '@delay: 推迟时间
    
    Dim found As Range
    Dim col, content As String
    Dim index, lastRow, tsOffset As Integer
    
    
    Set found = Sheets("语音提示").Range("A:A").Find(target, , , xlWhole)
    col = found.Offset(0, cOffset).value
    
    With Sheets("语音提示")
        '找到语音提醒的最后一行
        lastRow = .Range(col & "1", .Range(col & "1").End(xlDown)).Rows.Count
        'Debug.Print lastRow
        'Debug.Print tAnchor
        
        '语音提醒是从第二行开始
        For index = 2 To lastRow
            'Debug.Print col & ": " & .Range(col & index).value
            
            '语音提醒时间与锚点的时间间隔
            tsOffset = .Range(col & index).value + delay
            content = .Range(col & index).Offset(0, 1).value
            speakLater tAnchor + TimeSerial(0, tsOffset, 0), content
            showMsgLater tAnchor + TimeSerial(0, tsOffset, 0), content
        Next index
    
    End With
       
End Sub





