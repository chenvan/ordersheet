Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(8, 9)
    
    If Not IsArray(target.value) Then
    
        Dim endTimeColumn As Integer
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
        
        
        
        If target.Column = endTimeColumn And Util.IsNum(target.value) Then
           
            
            Dim serialNum, nextStartTime, nextTips As Range
            Dim tipsOfSwitchCutting, tipsOfNow As String
            Dim tsOfSwitchCutting, tsOfNextReady As Integer
            Dim tpOfSwitchCutting, tpOfNextReady As Variant
            Dim currentTobacco, nextTobacco As String
            
            
'            currentTobacco = target.offset(0, 3 - target.Column).value
'            nextTobacco = target.offset(1, 3 - target.Column).value
'
'            Util.shedule currentTobacco, target.value, 3, False
            
            Set serialNum = target.Offset(0, -7)
            Set nextStartTime = target.Offset(1, -1)
            Set nextTips = target.Offset(1, 2)
            
            tipsOfSwitchCutting = "六四切丝换工单，留意叶柜出柜状态"
            tipsOfNow = "有回掺需通知膨梗中控，HDT工"
            
            If serialNum.value = "4" Or serialNum.value = "8" Then
                tsOfSwitchCutting = 9
                tsOfNextReady = 24
            Else:
                tsOfSwitchCutting = 5
                tsOfNextReady = 20
            End If
            
            nextStartTime.value = target.value + TimeSerial(0, tsOfNextReady, 0)
            
            Util.speakLater Now, tipsOfNow
            Util.showMsgLater Now, tipsOfNow
            
            tpOfSwitchCutting = target.value + TimeSerial(0, tsOfSwitchCutting, 0)
            
            If Time > tpOfSwitchCutting Then
                Util.speakLater Now, tipsOfSwitchCutting
                Util.showMsgLater Now, tipsOfSwitchCutting
            Else:
                Util.speakLater tpOfSwitchCutting, tipsOfSwitchCutting
                Util.showMsgLater tpOfSwitchCutting, tipsOfSwitchCutting
            End If
            
            tpOfNextReady = target.value + TimeSerial(0, tsOfNextReady, 0)
            'Debug.Print nextTips
            'Debug.Print nextTips.value
            If Not IsError(nextTips) Then
                'Debug.Print Time
                'Debug.Print tpOfNextReady
                If Time > tpOfNextReady Then
                    Util.speakLater Now, nextTips.value
                    Util.showMsgLater Now, nextTips.value
                Else:
                    Util.speakLater tpOfNextReady, nextTips.value
                    Util.showMsgLater tpOfNextReady, nextTips.value
                End If
            End If
        End If
    
    End If
    
End Sub