Private Sub Worksheet_Change(ByVal target As range)
    Dim tobaccoName As String
    '判断发生值变化的是前4行，第1行不为空值，第2到4行为数值
    If Util.IsInArray(target.Row, Array(1, 2, 3, 4)) And target.Column = 3 And isRangeValid(range("C2", "C4")) And range("C1") <> "" Then
        
        
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
        tobaccoName = range("C1").value
        vol = Util.getParam(tobaccoName, "烘丝秤流量")
        sxWaterAdd = Util.getParam(tobaccoName, "SX水分增加")
        stdWaterRemove = Util.getParam(tobaccoName, "标准除水量")
        ck1 = Util.getParam(tobaccoName, "ck1")

        
        range("C5").value = vol
        range("C6").value = sxWaterAdd
        range("C7").value = stdWaterRemove
        range("C8").value = ck1
        
        
        '拿到输入值
        inputWaterContent = range("C2").value
        outputWaterContent = range("C3").value
        realTemp = range("C4").value
        
        
        dryWeight = vol * (1 - inputWaterContent * 0.01)
        inputWeight = dryWeight / (1 - (inputWaterContent + sxWaterAdd) * 0.01)
        outputWeight = dryWeight / (1 - outputWaterContent * 0.01)
        realWaterRemove = inputWeight - outputWeight
        
        settingTemp = (stdWaterRemove - realWaterRemove) * ck1 + realTemp
        
        range("C9").value = dryWeight
        range("C10").value = inputWeight
        range("C11").value = outputWeight
        range("C12").value = realWaterRemove
        range("C13").value = settingTemp

    End If
    
    If Util.IsInArray(target.Row, Array(1, 2)) And target.Column = 8 And isRangeValid(range("H2")) And range("H1") <> "" Then
        Dim Pt As Double
        Dim Gt As Double
        Dim Ct As Double
        Dim ratio As Double
        
        tobaccoName = range("H1").value
        Pt = Util.getParam(tobaccoName, "膨丝回掺量")
        Gt = Util.getParam(tobaccoName, "梗丝回掺量")
        
        range("H3").value = Pt
        range("H4").value = Gt
        Ct = range("H2").value
        
        ratio = (Ct + Ct * 0.0291) * 100 / (10000 - Pt - Gt)
        
        range("H5") = ratio
        
    End If
End Sub

Private Function isRangeValid(ByRef target As range) As Boolean
    Dim cell As range
    For Each cell In target
        If Not Util.IsNum(cell.value) Then
            isRangeValid = False
            Exit Function
        End If
    Next cell
    isRangeValid = True
End Function

