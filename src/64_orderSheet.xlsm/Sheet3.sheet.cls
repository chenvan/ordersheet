Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(7, 8)
    
    If Not IsArray(target.value) Then
        Dim beginTimeColumn, endTimeColumn As Integer
        Dim currentTobacco, nextTobacco As String
        Dim dateOfSelectedRow As Variant

        beginTimeColumn = 7
        endTimeColumn = 8
        
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
            Dim storeName As String
            
            currentTobacco = target.Offset(0, 3 - target.Column).value
            dateOfSelectedRow = target.Offset(0, 1 - target.Column).value
            storeName = target.Offset(0, 6 - target.Column).value
            
            'Debug.Print (currentTobacco)
            'Debug.Print dateOfSelectedRow
            'Debug.Print "贮叶柜: " & storeName
            
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Or IsEmpty(dateOfSelectedRow) Then
                
                If storeName = "" Then
                    Util.speakAsync "补全储叶柜, 并重新填写加料开始时间"
                    Exit Sub
                End If
                Util.sheduleVoiceTips "加料", currentTobacco, "开始后", target.value, 0
                '入柜提醒
                Util.sheduleVoiceTipsAboutStore "贮叶柜", currentTobacco, storeName, target.value, 0
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
            
            currentTobacco = target.Offset(0, 3 - target.Column).value
            nextTobacco = target.Offset(1, 3 - target.Column).value
            
            If nextTobacco = "" Then
                '提醒写下一批次的烟牌,并重新填写这一批次的结束时间
                'Debug.Print "空白牌号"
                Util.speakAsync "填下一批牌号, 并重新填写加料结束时间"
                Exit Sub
            End If
            
            dateOfSelectedRow = target.Offset(0, 1 - target.Column).value
            
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Or IsEmpty(dateOfSelectedRow) Then
                If currentTobacco <> nextTobacco Then
                    '下一批次为不同牌号需要洗加料管, 延时10分钟
                    feedLiquidSweepDelay = 10
                End If
                
                '批结束提醒
                Util.sheduleVoiceTips "加料", currentTobacco, "结束后", target.value, 0
                '下一批开始前提醒
                Util.sheduleVoiceTips "加料", nextTobacco, "开始前", target.value, feedLiquidSweepDelay
            Else
                 Util.showMsg "日期并非今天, 加料结束时间不会触发语音提醒"
            End If
            
        End If
    End If
End Sub

Public Sub runFirstBatchTip()
   Util.runFirstBatchTip "加料"
End Sub
