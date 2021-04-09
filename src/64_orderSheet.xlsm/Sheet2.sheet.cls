Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(6, 7)
    
    If Not IsArray(target.value) Then
        Dim beginTimeColumn, endTimeColumn As Integer
        Dim currentTobacco, nextTobacco As String
        Dim dateOfSelectedRow As Variant
        Dim oldTargetValue As Variant
        
        beginTimeColumn = 6
        endTimeColumn = 7
        
        dateOfSelectedRow = target.Offset(0, 1 - target.Column).value
        
        If Util.IsInArray(target.Column, timeFormatColumns) And IsNumeric(target.value) Then
            'inputTime's column is time format, which is a double type
            'we input an int, but it will change to double type
            
            If Util.IsIntType(target.value) Then
                Application.EnableEvents = False
                
                Application.Undo
                oldTargetValue = target.value
                Application.Undo
                
                If target.value <> "" Then
                    target.value = Util.ParseInputTime(target.value, dateOfSelectedRow)
                End If
                
                Application.EnableEvents = True
            End If
            
        End If
        
        If target.Column = beginTimeColumn And IsNumeric(target.value) Then
            '开始时间列触发批开始语音提醒
            
            currentTobacco = target.Offset(0, 3 - target.Column).value
            
   
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Then
                If oldTargetValue <> "" Then
                    '取消用上一时间锚定的提醒
                    Util.showMsg "取消以" & Format(oldTargetValue, "hh:mm") & "为锚定的语音触发"
                    Util.sheduleVoiceTips "回潮", currentTobacco, "开始后", oldTargetValue, 0, False
                End If
                
                If target.value <> "" Then
                    Util.sheduleVoiceTips "回潮", currentTobacco, "开始后", target.value, 0
                End If
            Else
                Util.showMsg "日期并非今天, 回潮开始时间不会触发语音提醒"
            End If
        
        End If
        
        If target.Column = endTimeColumn And IsNumeric(target.value) Then
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
                If oldTargetValue <> "" Then
                    '取消用上一次时间锚定的批结束提醒
                    Util.showMsg "取消以" & Format(oldTargetValue, "hh:mm") & "为锚定的语音触发"
                    Util.sheduleVoiceTips "回潮", currentTobacco, "结束后", oldTargetValue, 0, False
                    '取消用上一次时间锚定的下一批开始前提醒
                    Util.sheduleVoiceTips "回潮", nextTobacco, "开始前", oldTargetValue, 0, False
                End If
                
                If target.value <> "" Then
                    '批结束提醒
                    Util.sheduleVoiceTips "回潮", currentTobacco, "结束后", target.value, 0
                    '下一批开始前提醒
                    Util.sheduleVoiceTips "回潮", nextTobacco, "开始前", target.value, 0
                End If
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
    Dim output As String
    
    output = Util.guessFinishTime("回潮段")
    MsgBox output
End Sub

