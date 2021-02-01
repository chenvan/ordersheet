Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(6, 7)
    
    If Not IsArray(target.value) Then
        Dim beginTimeColumn, endTimeColumn As Integer
        Dim currentTobacco, nextTobacco As String
        Dim dateOfSelectedRow As Variant
        
        beginTimeColumn = 6
        endTimeColumn = 7
        
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
        
        If target.Column = beginTimeColumn And Util.IsNum(target.value) Then
            '开始时间列触发批开始语音提醒
            
            currentTobacco = target.Offset(0, 3 - target.Column).value
            
   
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Then
                Util.sheduleVoiceTips "回潮", currentTobacco, "开始后", target.value, 0
            Else
                Util.showMsg "日期并非今天, 回潮开始时间不会触发语音提醒"
            End If
        
        End If
        
        If target.Column = endTimeColumn And Util.IsNum(target.value) Then
            '结束时间列触发批结束语音提醒, 下一批开始前提醒语音
            
            '注意下一批次是空白的情况
            
            currentTobacco = target.Offset(0, 3 - target.Column).value
            nextTobacco = target.Offset(1, 3 - target.Column).value

            If nextTobacco = "" Then
                '提醒写下一批次的烟牌,并重新填写这一批次的结束时间
                'Debug.Print "空白牌号"
                Util.speakAsync "请填下一批的牌号, 并重新填写这批回潮结束时间"
                Exit Sub
            End If
            
            dateOfSelectedRow = target.Offset(0, 1 - target.Column).value
            
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Then
                '批结束提醒
                Util.sheduleVoiceTips "回潮", currentTobacco, "结束后", target.value, 0
                '下一批开始前提醒
                Util.sheduleVoiceTips "回潮", nextTobacco, "开始前", target.value, 0
            Else
                 Util.showMsg "日期并非今天, 回潮结束时间不会触发语音提醒"
            End If
        End If
    End If
End Sub

Public Sub checkBeforeWork()
   Util.checkBeforeWork "回潮"
End Sub

Public Sub guessFinishTime()
'    Dim endTimeColumn, tobaccoNameColumn As Integer
'    Dim tobaccoNameLastRow, endTimeLastRow As Integer
'    Dim switchTime, tobaccoFinishTime As Integer
'    Dim index As Integer
'    Dim beginTime, endTime As Date
'    Dim output, tobaccoName As String
'
'    'beginTimeColumn = Util.getColumnIndex("回潮段", "回潮秤开始时间")
'    endTimeColumn = Util.getColumnIndex("回潮段", "结束时间")
'    tobaccoNameColumn = Util.getColumnIndex("回潮段", "牌号")
'    switchTime = Sheets("设定").Range("A:A").Find("回潮段转烟时间").Offset(0, 1).value
'
'    tobaccoNameLastRow = Cells(Rows.count, tobaccoNameColumn).End(xlUp).Row + 1
'    endTimeLastRow = Cells(Rows.count, endTimeColumn).End(xlUp).Row + 1
'
'    If Cells(endTimeLastRow, endTimeColumn - 1).value = "" Then
'        beginTime = DateAdd("n", switchTime, Cells(endTimeLastRow - 1, endTimeColumn).value)
'    Else
'        beginTime = Cells(endTimeLastRow, endTimeColumn - 1).value
'    End If
'
'
'    For index = endTimeLastRow To tobaccoNameLastRow - 1
'        tobaccoName = Cells(index, tobaccoNameColumn).value
'        tobaccoFinishTime = Util.getParam(tobaccoName, "回潮段生产时长")
'
'        endTime = DateAdd("n", tobaccoFinishTime, beginTime)
'
'        output = output + tobaccoName + " : " + Format(beginTime, "hh:mm") + " - " + Format(endTime, "hh:mm") + vbCrLf
'
'
'        beginTime = DateAdd("n", switchTime, endTime)
'    Next
'
'    MsgBox output
    Dim output As String
    
    output = Util.guessFinishTime("回潮段")
    MsgBox output
End Sub

