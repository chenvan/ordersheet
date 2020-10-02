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
    Application.OnTime laterTime, "'speakAsync """ & content & "'"
End Sub

Sub speakAsync(ByVal content As String)
    Application.Speech.speak content & "��" & content, True
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
    
    'ͨ��sheetName, producePhase��������������б�
    Set tipArray = loadTips(sheetName, producePhase)
    '�����б�, ����Ƿ���Ҫ�ı�ʱ��
    
    For Each tipSet In tipArray
        '��������Ƿ���Ҫ����
        If tipSet.Exists("redirect") Then
            realContent = genNewContent(tobaccoName, tipSet("redirect"))
        ElseIf tipSet.Exists("hdt") Then
            '����Ƿ���Ҫ�ز�HDT
            If getParam(tobaccoName, "HDT�������") = "" Then GoTo Continue
            realContent = tipSet("hdt")
        Else
            realContent = tipSet("content")
        End If
        
        '���ʱ���Ƿ���Ҫ����
        If tipSet.Exists("aOffsetTime") Then
            realOffsetTime = adjustOffsetTime(tobaccoName, tipSet("aOffsetTime"))
        Else
            realOffsetTime = tipSet("sOffsetTime")
        End If
        
        'Debug.Print realContent
        
        'װ������
        pushVoiceTip realContent, realOffsetTime + delay, baseTime, tipSet("isForceBroadcast")
       
Continue:
    Next tipSet
    
End Sub

Sub sheduleVoiceTipsAboutStore(ByVal storePlace As String, ByVal tobaccoName As String, ByVal storeName As String, ByVal baseTime As Variant, ByVal delay As Integer)
    Dim tipArray As Variant
    Dim tipSet As Scripting.Dictionary
    Dim realOffsetTime As Variant
    Dim realStore As String
    
    If storePlace = "��Ҷ��" Then
        '���ؼ������tips
        Dim storeNamesInfeedLiquid As Variant
    
        storesInFeedLiquid = Array("1", "2", "3", "4", "5", "6")
        realStore = "HDT"
        
        If IsInArray(storeName, storesInFeedLiquid) Then
            realStore = "Ҷ��"
        End If
        
    Else
        '�����к����tips
        Dim storesInAddEssence As Variant
        Dim storeIndex As Integer
        
        storesInAddEssence = Array("��A", "��B", "��C", "��D", "��E", "��F", "��G", "��H", "��J", "��K", "��L", "��M", "��N", "��P", "��Q", "��R", "��S", "��T", "ľA", "ľB", "ľC", "��A", "��B", "��C", "��D", "��E", "��F", "��G", "��H", "��A", "��B")
        storeIndex = Application.Match(storeName, storesInAddEssence, False)
        
        realStore = "��AF"
        
        If storeIndex > 6 And storeIndex < 13 Then
            realStore = "��GM"
        ElseIf storeIndex > 12 And storeIndex <= 18 Then
            realStore = "��NT"
        ElseIf storeIndex > 18 And storeIndex <= 21 Then
            realStore = "ľAC"
        ElseIf storeIndex > 21 And storeIndex <= 25 Then
            realStore = "��AD"
        ElseIf storeIndex > 25 And storeIndex <= 29 Then
            realStore = "��EH"
        ElseIf storeIndex > 29 And storeIndex <= 31 Then
            realStore = "��AB"
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
        showMsg "��ʱ,��������������: $" & content
    End If
End Sub

Sub runFirstBatchTip(ByVal sheetName As String)
    Dim found As range
    Dim firstTobaccoName As String
    'find today's first row
    Set found = Sheets(sheetName).range("A:A").Find(Date)
    
    If found Is Nothing Then
        Util.showMsg "û���ҵ����������"
    Else
        firstTobaccoName = found.offset(0, 2).value
        'ʱ��ʹ��now�Ļ���ʱ��ת�������
        sheduleVoiceTips sheetName, firstTobaccoName, "��һ��", Time, 0
    End If
End Sub

Public Function loadTips(ByVal fLayerP As String, ByVal sLayerP As String) As Variant
    Dim fullPath As String
    Dim allTips As Scripting.Dictionary
    
    fullPath = Sheets("�趨").range("A:A").Find("�����ļ�·��").offset(0, 1).value
    
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
    
    rowIndex = Sheets("�趨").range("A2:A18").Find(tobaccoName).Row
    columnIndex = Sheets("�趨").range("A1:Z1").Find(paramName).Column
    
    getParam = Sheets("�趨").Cells(rowIndex, columnIndex)
End Function

Function genNewContent(ByVal tobaccoName As String, ByVal paramName As String) As String
    Dim param As String
    
    param = getParam(tobaccoName, paramName)
    
    genNewContent = paramName & param

End Function


Function adjustOffsetTime(ByVal tobaccoName As String, ByVal offsetTime As Integer) As Integer
    Dim mainTobaccoVolume As Integer
    
    mainTobaccoVolume = getParam(tobaccoName, "��Ҷ˿������")
    
    adjustOffsetTime = offsetTime * 6250 / mainTobaccoVolume
End Function



