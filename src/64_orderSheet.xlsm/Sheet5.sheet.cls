Private Sub Worksheet_Change(ByVal target As Range)
    Dim tobaccoName As String
    '判断发生值变化的是前4行，第1行不为空值，第2到4行为数值
    If Util.IsInArray(target.Row, Array(1, 2, 3, 4)) And target.Column = 3 And isRangeValid(Range("C2", "C4")) And Range("C1") <> "" Then
        
        
        Dim vol As Double
        Dim ck1 As Double
        Dim stdWaterRemove As Double
        Dim sxWaterAdd As Double
        Dim dryWeight As Double
        Dim intputWeight As Double
        Dim outputWeight As Double
        Dim realWaterRemove As Double
        Dim outputWaterContent As Double
        Dim inputWaterContent As Double
        Dim realTemp As Double
        Dim settingTemp As Double
        
        '通过牌号拿到设定值
        tobaccoName = Range("C1").value
        vol = Util.getParam(tobaccoName, "烘丝秤流量")
        sxWaterAdd = Util.getParam(tobaccoName, "SX水分增加")
        stdWaterRemove = Util.getParam(tobaccoName, "标准除水量")
        ck1 = Util.getParam(tobaccoName, "ck1")

        
        Range("C5").value = vol
        Range("C6").value = sxWaterAdd
        Range("C7").value = stdWaterRemove
        Range("C8").value = ck1
        
        
        '拿到输入值
        inputWaterContent = Range("C2").value
        outputWaterContent = Range("C3").value
        realTemp = Range("C4").value
        
        
        dryWeight = vol * (1 - inputWaterContent * 0.01)
        inputWeight = dryWeight / (1 - (inputWaterContent + sxWaterAdd) * 0.01)
        outputWeight = dryWeight / (1 - outputWaterContent * 0.01)
        realWaterRemove = inputWeight - outputWeight
        
        settingTemp = (stdWaterRemove - realWaterRemove) * ck1 + realTemp
        
        Range("C9").value = dryWeight
        Range("C10").value = inputWeight
        Range("C11").value = outputWeight
        Range("C12").value = realWaterRemove
        Range("C13").value = settingTemp

    End If
    
    If Util.IsInArray(target.Row, Array(1, 2)) And target.Column = 8 And isRangeValid(Range("H2")) And Range("H1") <> "" Then
        Dim Pt, Gt, Ct, CSettingRatio, CDiff, realRatio As Double
        
        tobaccoName = Range("H1").value
        Pt = Util.getParam(tobaccoName, "膨丝回掺量")
        Gt = Util.getParam(tobaccoName, "梗丝回掺量")
        CSettingRatio = Util.getParam(tobaccoName, "残丝回掺比例")
        CDiff = Util.getParam(tobaccoName, "残丝批间差异系数")
        
        Range("H3").value = Pt
        Range("H4").value = Gt
        Range("H5").value = CDiff
        Range("H6").value = CSettingRatio
        Ct = Range("H2").value
        
        realRatio = (Ct + Ct * CSettingRatio * CDiff) * 100 / (10000 - Pt - Gt)
        
        Range("H7") = realRatio
        
    End If
End Sub

Private Function isRangeValid(ByRef target As Range) As Boolean
    Dim cell As Range
    For Each cell In target
        If Not Util.IsNum(cell.value) Then
            isRangeValid = False
            Exit Function
        End If
    Next cell
    isRangeValid = True
End Function


