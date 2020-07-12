Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(8, 9)
    
    If Not IsArray(target.value) Then
        Dim beginTimeColumn, endTimeColumn As Integer
        Dim currentTobacco, nextTobacco As String
        
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
            
            currentTobacco = target.Offset(0, 3 - target.Column).value
            
            'Debug.Print (currentTobacco)
            
            Util.shedule currentTobacco, target.value, 3, False
        
        End If
        
        If target.Column = endTimeColumn And Util.IsNum(target.value) Then
            '结束时间列触发批结束语音提醒
            
            '通过对比下一批次与现批次的牌号判断加料是否需要清扫
            '注意下一批次是空白的情况
            
            '加料是否需要清洁加料管
            Dim isFeedLiquidSweep As Boolean
            
            isFeedLiquidSweep = False
            currentTobacco = target.Offset(0, 3 - target.Column).value
            nextTobacco = target.Offset(1, 3 - target.Column).value
            
            If currentTobacco <> nextTobacco And nextTobacco <> "" Then
                isFeedLiquidSweep = True
            ElseIf nextTobacco = "" Then
                '提醒写下一批次的烟牌,并重新填写这一批次的结束时间
                '中断sub
                Debug.Print "空白牌号"
                Exit Sub
            End If
            
            '批结束的提醒是通用的提醒, 烟牌号用现在的和下一批的没有区别
            Util.shedule nextTobacco, target.value, 4, isFeedLiquidSweep
        End If
    End If
End Sub