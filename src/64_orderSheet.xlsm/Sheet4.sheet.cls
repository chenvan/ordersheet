Private Sub Worksheet_Change(ByVal target As range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(8, 9)
    
    If Not IsArray(target.value) Then
    
        Dim currentTobacco, nextTobacco As String
        Dim dateOfSelectedRow As Variant
        Dim beginTimeColumn, endTimeColumn As Integer
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
                Util.shedule currentTobacco, target.value, 8, 0
            Else
                Util.showMsg "日期并非今天, 加料开始时间不会触发语音提醒"
            End If
            
        End If
        
        If target.Column = endTimeColumn And Util.IsNum(target.value) Then
           
            '结束时间
            '不再自动填写下一批的开始时间
            Dim serialNum As range
            Set serialNum = target.offset(0, -7)
            
            Dim addEssenceSweepDelay As Integer
            addEssenceSweepDelay = 0
            
            currentTobacco = target.offset(0, 3 - target.Column).value
            nextTobacco = target.offset(1, 3 - target.Column).value
            
            If nextTobacco = "" Then
                '提醒写下一批次的烟牌,并重新填写这一批次的结束时间
                'Debug.Print "空白牌号"
                Util.speakAsync "请填下一批的牌号, 并重新填写这批烘丝结束时间"
                Exit Sub
            End If
            
            dateOfSelectedRow = target.offset(0, 1 - target.Column).value
            
            '行的日期为今天或者是空内容时才进行提醒
            If dateOfSelectedRow = Date Or IsEmpty(dateOfSelectedRow) Then
                '第4批和第8批加香机需要清扫, 延时4分钟上烟
                If serialNum.value = "4" Or serialNum.value = "8" Then
                    addEssenceSweepDelay = 4
                End If
                
                '批结束提醒
                Util.shedule currentTobacco, target.value, 9, 0
                '下一批开始前提醒
                Util.shedule nextTobacco, target.value, 7, addEssenceSweepDelay
                '下一批次特殊提醒
                Util.speakPrecaution nextTobacco, target.value, "O", addEssenceSweepDelay
            Else
                 Util.showMsg "日期并非今天, 加料结束时间不会触发语音提醒"
            End If
        End If
    End If
End Sub

Public Sub firstBatchTip()
'    Dim found As range
'    Dim firstTobaccoName As String
'    'find today's first row
'    Set found = Sheets("切烘加香段").range("A:A").Find(Date, , , xlWhole)
'    If found Is Nothing Then
'        Util.showMsg "没有找到今天的日期"
'    Else
'        'get the tobacco name, run util.speakPrecation & 开始前提醒
'        firstTobaccoName = found.offset(0, 2).value
'        'Debug.Print firstTobaccoName
'        '需要尽快提醒, 所以把delay设为负数
'        '时间使用now的话会时间转换会溢出
'        Util.shedule firstTobaccoName, Time, 7, -10
'        Util.speakPrecaution firstTobaccoName, Time, "O", -15
'    End If
    
    Util.runFirstBatchWarning "切烘加香段", 7, "O"
End Sub

