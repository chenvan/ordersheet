Private Sub Worksheet_Change(ByVal target As range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(8, 9)
    
    If Not IsArray(target.value) Then
        Dim beginTimeColumn, endTimeColumn As Integer
        Dim currentTobacco, nextTobacco As String
        Dim dateOfSelectedRow As Variant
        
        beginTimeColumn = 8
        endTimeColumn = 9
        
        If Util.IsInArray(target.Column, timeFormatColumns) And Util.IsNum(target.value) Then
            'inputTime's column is time format, which is a double type
            'we input an int, but it will change to double type
            
            If Util.IsIntType(target.value) Then
                Application.EnableEvents = False
                target.value = Others.ParseInputTime(target.value)
                Application.EnableEvents = True
            End If
            
        End If
        
        If target.Column = beginTimeColumn And Util.IsNum(target.value) Then
            '开始时间列触发批开始语音提醒
            
            currentTobacco = target.offset(0, 3 - target.Column).value
            dateOfSelectedRow = target.offset(0, 1 - target.Column).value
            
            'Debug.Print (currentTobacco)
            'Debug.Print dateOfSelectedRow
            
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Or IsEmpty(dateOfSelectedRow) Then
                Util.shedule currentTobacco, target.value, 5, 0
            Else
                Util.showMsg "日期并非今天, 加料开始时间不会触发语音提醒"
            End If
            
        End If
        
        If target.Column = endTimeColumn And Util.IsNum(target.value) Then
            '结束时间列触发批结束语音提醒, 下一批开始前提醒语音
            
            '通过对比下一批次与现批次的牌号判断加料是否需要延时
            '注意下一批次是空白的情况
            
            Dim feedLiquidSweepDelay As Integer
            feedLiquidSweepDelay = 0
            
            currentTobacco = target.offset(0, 3 - target.Column).value
            nextTobacco = target.offset(1, 3 - target.Column).value
            
            If nextTobacco = "" Then
                '提醒写下一批次的烟牌,并重新填写这一批次的结束时间
                'Debug.Print "空白牌号"
                Util.speakAsync "请填下一批的牌号, 并重新填写这批加料结束时间"
                Exit Sub
            End If
            
            dateOfSelectedRow = target.offset(0, 1 - target.Column).value
            
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Or IsEmpty(dateOfSelectedRow) Then
                If currentTobacco <> nextTobacco Then
                    '下一批次为不同牌号需要洗加料管, 延时10分钟
                    feedLiquidSweepDelay = 10
                End If
                
                '批结束提醒
                Util.shedule currentTobacco, target.value, 6, 0
                '下一批开始前提醒
                Util.shedule nextTobacco, target.value, 4, feedLiquidSweepDelay
            Else
                 Util.showMsg "日期并非今天, 加料结束时间不会触发语音提醒"
            End If
            
        End If
    End If
End Sub
