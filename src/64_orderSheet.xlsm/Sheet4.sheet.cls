Option Explicit

Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    Dim beginTimeColumn, endTimeColumn As Integer

    beginTimeColumn = Util.getColumnIndex("切烘加香段", "开始时间")
    endTimeColumn = Util.getColumnIndex("切烘加香段", "结束时间")
    
    timeFormatColumns = Array(beginTimeColumn, endTimeColumn)
    
    If Not IsArray(target.value) Then
    
        Dim dateOfSelectedRow As Variant
        Dim oldTargetValue As Variant
        
        dateOfSelectedRow = Util.getValue("切烘加香段", "日期", target.Row)
        
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
            
            '行的日期为今天时才进行提醒
            If dateOfSelectedRow = Date Then
                Util.triggerBeginTimeVoiceTips "切烘加香段", oldTargetValue, target
            Else
                Util.showMsg "日期并非今天, 烘丝开始时间不会触发语音提醒"
            End If
            
        End If
        
        'IsNumeric function will let empty string be true
        If target.Column = endTimeColumn And IsNumeric(target.value) Then
           
            '结束时间
            
            '行的日期为今天才进行提醒
            If dateOfSelectedRow = Date Then
                Util.triggerEndTimeVoiceTips "切烘加香段", oldTargetValue, target
            Else
                Util.showMsg "日期并非今天, 烘丝结束时间不会触发语音提醒"
            End If
        End If
    End If
End Sub

Public Sub checkBeforeWork()
   Util.checkBeforeWork "切烘加香"
End Sub

Public Sub calcuRatio()
    Dim lastRow As Long
    Dim tobaccoName As String
    Dim Pt, Gt, Ct, CSettingRatio, CDiff, realRatio As Double
    Dim endTimeColumn, tobaccoNameColumn As Integer
    
    endTimeColumn = Util.getColumnIndex("切烘加香段", "结束时间")
    tobaccoNameColumn = Util.getColumnIndex("切烘加香段", "牌号")
    
    lastRow = Cells(Rows.count, endTimeColumn).End(xlUp).Row + 1
    tobaccoName = Cells(lastRow, tobaccoNameColumn)
    
    Ct = Application.InputBox("回掺残丝量: ", Title:=tobaccoName, Type:=1)
    
    If Ct Then
    
        Pt = Util.getParam(tobaccoName, "膨丝回掺量")
        Gt = Util.getParam(tobaccoName, "梗丝回掺量")
        CSettingRatio = Util.getParam(tobaccoName, "残丝回掺比例")
        CDiff = Util.getParam(tobaccoName, "残丝批间差异系数")
        
        realRatio = (Ct + Ct * CSettingRatio * CDiff) * 100 / (10000 - Pt - Gt)
        
        MsgBox "最大回掺比例: " & CSettingRatio & vbNewLine & "计算回掺比例: " & Format(realRatio, "0.000"), , tobaccoName
    End If
End Sub

Public Sub calcuSettingTemp(Optional ByVal ck1Src As String = "setting")
    '找出选中行的入口水分, 出口水分, 实际温度, 烟牌名
    Dim currentRow As Long
    Dim tobaccoName As String
    Dim inputWaterContent, outputWaterContent, realTemp, vol, sxWaterAdd, stdWaterRemove, ck1 As Double
    Dim dryWeight, inputWeight, outputWeight, realWaterRemove, settingTemp As Double
    Dim tobaccoColIndex, inputWaterColIndex, outputWaterColIndex, realTempColIndex As Integer
    
    currentRow = ActiveCell.Row
    tobaccoColIndex = Util.getColumnIndex("切烘加香段", "牌号")
    inputWaterColIndex = Util.getColumnIndex("切烘加香段", "来料水分")
    outputWaterColIndex = Util.getColumnIndex("切烘加香段", "出口水分设定")
    realTempColIndex = Util.getColumnIndex("切烘加香段", "TT1温度实际")
    
    
    tobaccoName = Cells(currentRow, tobaccoColIndex)
    inputWaterContent = Cells(currentRow, inputWaterColIndex)
    outputWaterContent = Cells(currentRow, outputWaterColIndex)
    realTemp = Cells(currentRow, realTempColIndex)
    
    If IsEmpty(inputWaterContent) Or IsEmpty(outputWaterContent) Or IsEmpty(tobaccoName) Or IsEmpty(realTemp) Then
        Exit Sub
    End If
    
    vol = Util.getParam(tobaccoName, "烘丝秤流量")
    sxWaterAdd = Util.getParam(tobaccoName, "SX水分增加")
    stdWaterRemove = Util.getParam(tobaccoName, "标准除水量")
    
    If ck1Src = "setting" Then
        ck1 = Util.getParam(tobaccoName, "ck1")
    ElseIf ck1Src = "new" Then
        ck1 = Util.getParam(tobaccoName, "新ck1")
    End If
    
    dryWeight = vol * (1 - inputWaterContent * 0.01)
    inputWeight = dryWeight / (1 - (inputWaterContent + sxWaterAdd) * 0.01)
    outputWeight = dryWeight / (1 - outputWaterContent * 0.01)
    realWaterRemove = Round(inputWeight - outputWeight)
    
    settingTemp = (stdWaterRemove - realWaterRemove) * ck1 + realTemp
    
    MsgBox "烘丝秤流量: " & vol & vbCrLf _
            & "SX水分增加: " & sxWaterAdd & vbCrLf _
            & "ck1: " & Format(ck1, "0.000") & vbCrLf & vbCrLf _
            & "使用标准除水量计算温度: " & vbCrLf _
            & "标准除水量: " & stdWaterRemove & vbCrLf _
            & "设定温度: " & Format(settingTemp, "000.0") & vbCrLf & vbCrLf _
            & "修改实际除水量为标准除水量: " & vbCrLf _
            & "标准除水量: " & realWaterRemove & vbCrLf _
            & "设定温度: " & Format(realTemp, "000.0"), , tobaccoName
End Sub


Public Sub calcuSettingTempWithNewCk1()
    calcuSettingTemp "new"
End Sub


Public Sub guessFinishTime()
    Dim output As String
    
    output = Util.guessFinishTime("切烘加香段")

    MsgBox output
End Sub
