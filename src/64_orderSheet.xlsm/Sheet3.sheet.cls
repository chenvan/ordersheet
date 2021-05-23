Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    Dim beginTimeColumn, endTimeColumn As Integer
    
    beginTimeColumn = Util.getColumnIndex("加料段", "开始时间")
    endTimeColumn = Util.getColumnIndex("加料段", "结束时间")
    
    timeFormatColumns = Array(beginTimeColumn, endTimeColumn)
    
    If Not IsArray(target.value) Then
        Dim dateOfSelectedRow As Variant
        Dim oldTargetValue As Variant
        
        dateOfSelectedRow = Util.getValue("加料段", "日期", target.Row)
         
         'IsNumeric function will let empty string be true
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
        
        'IsNumeric function will let empty string be true
        If target.Column = beginTimeColumn And IsNumeric(target.value) Then
            '开始时间列触发批开始语音提醒
          
            
            '行的日期为今天才进行提醒
            If dateOfSelectedRow = Date Then
                Util.triggerBeginTimeVoiceTips "加料段", oldTargetValue, target
            Else
                Util.showMsg "日期并非今天, 加料开始时间不会触发语音提醒"
            End If
            
        End If
        
        If target.Column = endTimeColumn And IsNumeric(target.value) Then
            '结束时间列触发批结束语音提醒, 下一批开始前提醒语音
            
            '通过对比下一批次与现批次的牌号判断加料是否需要延时
            '注意下一批次是空白的情况
            
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Then
                Util.triggerEndTimeVoiceTips "加料段", oldTargetValue, target
            Else
                Util.showMsg "日期并非今天, 加料结束时间不会触发语音提醒"
            End If
            
        End If
    End If
End Sub

Public Sub checkBeforeWork()
   Util.checkBeforeWork "加料段"
End Sub

Public Sub guessFinishTime()
    Dim output As String
    
    output = Util.guessFinishTime("加料段")

    MsgBox output
End Sub

