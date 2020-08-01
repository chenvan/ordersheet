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
    
    answer = MsgBox("�˲������������������գ��Ƿ��Ѿ����ļ����", vbYesNoCancel, "����")
    
    If answer = 6 Then
        Sheets("�س���").range("A3:A1171").ClearContents
        Sheets("�س���").range("C3:K1171").ClearContents
        Sheets("�س���").range("M3:N1171").ClearContents
        Sheets("�س���").range("P3:P1171").ClearContents
        
        Sheets("���϶�").range("A3:A1321").ClearContents
        Sheets("���϶�").range("C3:D1321").ClearContents
        Sheets("���϶�").range("G3:K1321").ClearContents
        Sheets("���϶�").range("N3:P1321").ClearContents
        Sheets("���϶�").range("R3:R1321").ClearContents
        
        Sheets("�к�����").range("A3:A1251").ClearContents
        Sheets("�к�����").range("C3:D1251").ClearContents
        Sheets("�к�����").range("G3:J1251").ClearContents
        Sheets("�к�����").range("L3:AC1251").ClearContents
        Sheets("�к�����").range("AE3:AE1251").ClearContents
        
        Sheets("HDT��").range("A3:A152").ClearContents
        Sheets("HDT��").range("C3:D152").ClearContents
        Sheets("HDT��").range("L3:L152").ClearContents
        
    ElseIf answer = 7 Then
        varResult = Application.GetSaveAsFilename(filefilter:="Marco Enabled Workbook(*.xlsm), *xlsm")
        If varResult <> False Then
            ThisWorkbook.SaveAs (varResult)
        End If
    End If
End Sub

Sub speakLater(ByVal laterTime As Variant, ByVal content As String)
    '������ŵ�ʱ���Ѿ�����,Ӧ������������������
    'Debug.Print Time
    'Debug.Print laterTime
    If Time >= laterTime Then
        'Debug.Print "��ʱ"
        content = "��ʱ," & content
        Application.OnTime Now, "'speakAsync """ & content & "'"
    Else
        Application.OnTime laterTime, "'speakAsync """ & content & "'"
    End If
End Sub

Sub speakAsync(ByVal content As String)
    Application.Speech.speak content & "��" & content, True
End Sub

    
Sub showMsgLater(ByVal laterTime As Variant, ByVal content As String)
    If Time >= laterTime Then
        content = "��ʱ," & content
        Application.OnTime Now, "'showMsg """ & content & "'"
    Else
        Application.OnTime laterTime, "'showMsg """ & content & "'"
    End If
End Sub

Sub showMsg(ByVal content As String)
    Application.StatusBar = "##" & content & "   " & Left(Application.StatusBar, 80)
End Sub


Sub shedule(ByVal target As String, ByVal tAnchor As Variant, ByVal cOffset As Integer, ByVal delay As Integer, Optional ByVal isTriggerNow As Boolean = False)
    '������ʾ���ź���
    '@taget: ���ƺ�
    '@tAnchor: ʱ��ê��
    '@cOffset: ������������A�е��м��
    '@delay: �Ƴ�ʱ��
    '@isTriggerNow: ��������
    
    Dim found As range
    Dim col As String
    Dim index, lastRow As Integer
    Dim info As tipInfo
    
    Set found = Sheets("������ʾ").range("A:A").Find(target, , , xlWhole)
    col = found.offset(0, cOffset).value
    
    With Sheets("������ʾ")
        '�ҵ��������ѵ����һ��
        'use .End(xlDown).End(xlDown).End(xlUp) instead of only use .End(xlDown)
        'it can avoid return too many rowscount if the tips is empty
        lastRow = .range(col & "1", .range(col & "1").End(xlDown).End(xlDown).End(xlUp)).Rows.Count
        'Debug.Print lastRow
        'Debug.Print tAnchor
        
        '���������Ǵӵڶ��п�ʼ
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

    Set found = Sheets("������ʾ").range("A:A").Find(target, , , xlWhole)
    preCautionTip = getTipInfo(colName, found.Row, tAnchor, delay, isTriggerNow)

    runTipInfo preCautionTip, -11

End Sub

Private Function getTipInfo(ByVal colName As String, ByVal rowIndex As Integer, ByVal tAnchor As Variant, ByVal delay As Integer, ByVal isTriggerNow As Boolean) As tipInfo
    Dim info As tipInfo
    Dim tsOffset As Integer
    Dim triggerTime As Variant
    Dim content As String
    
    '��������ʱ����ê���ʱ����

    tsOffset = Sheets("������ʾ").range(colName & rowIndex).value + delay
    content = Sheets("������ʾ").range(colName & rowIndex).offset(0, 1).value
    
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
        showMsg "��ʱ����" & Abs(latestTime) - 1 & "����, ���������������, $" & info.tipContent
    End If
End Sub

Sub runFirstBatchWarning(ByVal sheetName, ByVal colNameForBeforeBegin, ByVal colNameForPrecaution)
    Dim found As range
    Dim firstTobaccoName As String
    'find today's first row
    Set found = Sheets(sheetName).range("A:A").Find(Date, , , xlWhole)
    If found Is Nothing Then
        Util.showMsg "û���ҵ����������"
    Else
        'get the tobacco name, run util.speakPrecation & ��ʼǰ����
        firstTobaccoName = found.offset(0, 2).value
        'Debug.Print firstTobaccoName
        'ʱ��ʹ��now�Ļ���ʱ��ת�������
        Util.shedule firstTobaccoName, Time, colNameForBeforeBegin, 0, True
        Util.speakPrecaution firstTobaccoName, Time, colNameForPrecaution, 0, True
    End If
End Sub









