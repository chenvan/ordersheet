Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(4)

    If Not IsArray(target.value) Then
        Dim dateOfSelectedRow As Variant
        
        dateOfSelectedRow = target.Offset(0, 1 - target.Column).value
        
        If Util.IsInArray(target.Column, timeFormatColumns) And Util.IsNum(target.value) Then
            'inputTime's column is time format, which is a double type
            'we input an int, but it will change to double type
            
            If Util.IsIntType(target.value) Then
                Application.EnableEvents = False
                target.value = Util.ParseInputTime(target.value, dateOfSelectedRow)
                Application.EnableEvents = True
            End If

        End If

    End If

End Sub

'Public Sub sync96HDT()
'    '打开96的文件
'    'On Error GoTo errHandler
'
'    Dim pathOf96Workbook, hdtSheetName As String
'    Dim lastRowOf64HdtSheet, index, lastRowOf96HdtSheet As Long
'    Dim hdtSheetOf96 As Worksheet
'    Dim workbookOf96 As Workbook
'
'    pathOf96Workbook = "C:\Users\AWang\code\orderSheet\96生产记录共享表.xlsm"
'    hdtSheetName = "HDT段"
'
'    Application.ScreenUpdating = False
'
'    Set workbookOf96 = Workbooks.Open(pathOf96Workbook, True, True)
'    Set hdtSheetOf96 = workbookOf96.Worksheets("HDT段")
'
'    '64HDT表最后一行
'    'lastRowOf64HdtSheet = range("A1", range("A1").End(xlDown).End(xlDown).End(xlUp)).Rows.Count
'    lastRowOf64HdtSheet = Cells(Worksheets("HDT段").Rows.Count, "A").End(xlUp).offset(1).Row
'    lastRowOf96HdtSheet = hdtSheetOf96.Cells(hdtSheetOf96.Rows.Count, "A").End(xlUp).offset(1).Row
'    Debug.Print "lastrow of 64: " & lastRowOf64HdtSheet
'    Debug.Print "lastrow of 96: " & lastRowOf96HdtSheet
'
'    '循环96HDT段的表
''     For index = 3 To lastRow
''
''            '语音提醒时间与锚点的时间间隔
''            tsOffset = .range(col & index).value + delay
''            triggerTime = tAnchor + TimeSerial(0, tsOffset, 0)
''            timeDiffInMin = (triggerTime - Time) * 1440
''            content = .range(col & index).Offset(0, 1).value
''
''            '检查时间, 如果没有超时10分钟, 就进行提醒
''            If timeDiffInMin > -11 Then
''                speakLater triggerTime, content
''                showMsgLater triggerTime, content
''            Else
''                showMsg "超时大于10分钟, 不会进行语音提醒, $" & content
''            End If
''        Next index
'
'
'    If Util.isInSheetRange("2007JDHDT#099", "加料段", "D:D") Then
'        Debug.Print "exist"
'    Else
'        Debug.Print "not exist"
'    End If
'
'
'
''errHandler:
'    workbookOf96.Close False
'    Application.ScreenUpdating = True
'End Sub

