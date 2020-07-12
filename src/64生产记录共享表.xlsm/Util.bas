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
    
    answer = MsgBox("�˲������������������գ��Ƿ��Ѿ����ļ����", vbYesNoCancel, "����")
    
    If answer = 6 Then
        Sheets("�س���").Range("A3:A1171").ClearContents
        Sheets("�س���").Range("C3:K1171").ClearContents
        Sheets("�س���").Range("M3:N1171").ClearContents
        Sheets("�س���").Range("P3:P1171").ClearContents
        
        Sheets("���϶�").Range("A3:A1321").ClearContents
        Sheets("���϶�").Range("C3:D1321").ClearContents
        Sheets("���϶�").Range("G3:K1321").ClearContents
        Sheets("���϶�").Range("N3:P1321").ClearContents
        Sheets("���϶�").Range("R3:R1321").ClearContents
        
        Sheets("�к�����").Range("A3:A1251").ClearContents
        Sheets("�к�����").Range("C3:D1251").ClearContents
        Sheets("�к�����").Range("G3:J1251").ClearContents
        Sheets("�к�����").Range("L3:AC1251").ClearContents
        Sheets("�к�����").Range("AE3:AE1251").ClearContents
        
        Sheets("HDT��").Range("A3:A152").ClearContents
        Sheets("HDT��").Range("C3:D152").ClearContents
        Sheets("HDT��").Range("L3:L152").ClearContents
        
    ElseIf answer = 7 Then
        varResult = Application.GetSaveAsFilename(filefilter:="Marco Enabled Workbook(*.xlsm), *xlsm")
        If varResult <> False Then
            ThisWorkbook.SaveAs (varResult)
        End If
    End If
End Sub

Sub speakLater(ByVal laterTime As Variant, ByVal content As String)
    '������ŵ�ʱ���Ѿ�����,Ӧ������������������
    
    Application.OnTime laterTime, "'speakAsync """ & content & "'"
End Sub

Sub speakAsync(ByVal content As String)
    Application.Speech.speak content & "��" & content, True
End Sub

Sub showMsgLater(ByVal laterTime As Variant, ByVal content As String)
     Application.OnTime laterTime, "'showMsg """ & content & "'"
End Sub

Sub showMsg(ByVal content As String)
    Application.StatusBar = "***" & content & "***"
End Sub


Sub shedule(ByVal target As String, ByVal tAnchor As Variant, ByVal cOffset As Integer, ByVal isSweep As Boolean)
    '������ʾ���ź���
    '@taget: ���ƺ�
    '@tAnchor: ʱ��ê��
    '@cOffset: ������������A�е��м��
    '@isSweep: �Ƿ���ɨ
    
    Dim found As Range
    Dim col, content As String
    Dim index, lastRow, tsOffset As Integer
    
    
    Set found = Sheets("������ʾ").Range("A:A").Find(target, , , xlWhole)
    col = found.Offset(0, cOffset).value
    
    If isSweep Then
        '���ֻ�ܼӵ�z
        col = Chr(Asc(col) + 2)
        'Debug.Print "��: " & col
    End If
    
    With Sheets("������ʾ")
        '�ҵ��������ѵ����һ��
        lastRow = .Range(col & "1", .Range(col & "1").End(xlDown)).Rows.Count
        'Debug.Print lastRow
        'Debug.Print tAnchor
        
        '���������Ǵӵڶ��п�ʼ
        For index = 2 To lastRow
            'Debug.Print col & ": " & .Range(col & index).value
            
            '��������ʱ����ê���ʱ����
            tsOffset = .Range(col & index).value
            content = .Range(col & index).Offset(0, 1).value
            speakLater tAnchor + TimeSerial(0, tsOffset, 0), content
            showMsgLater tAnchor + TimeSerial(0, tsOffset, 0), content
        Next index
    
    End With
       
End Sub





