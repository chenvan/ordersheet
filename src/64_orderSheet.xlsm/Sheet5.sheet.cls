Private Sub Worksheet_Change(ByVal target As Range)
    Dim found As Range
    
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
        Set found = Sheets("设定").Range("A:A").Find(Range("C1").value, , , xlWhole)
        Range("C5").value = found.Offset(0, 4).value
        Range("C6").value = found.Offset(0, 5).value
        Range("C7").value = found.Offset(0, 6).value
        Range("C8").value = found.Offset(0, 7).value

        
        '拿到输入值
        inputWaterContent = Range("C2").value
        outputWaterContent = Range("C3").value
        realTemp = Range("C4").value
        
        '拿到牌号设定值
        vol = Range("C5").value
        sxWaterAdd = Range("C6").value
        stdWaterRemove = Range("C7").value
        ck1 = Range("C8").value
        
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
        Dim Pt As Double
        Dim Gt As Double
        Dim Ct As Double
        Dim ratio As Double
        
        Set found = Sheets("设定").Range("A:A").Find(Range("H1").value, , , xlWhole)
        Range("H3").value = found.Offset(0, 8).value
        Range("H4").value = found.Offset(0, 9).value
        
        Pt = Range("H3").value
        Gt = Range("H4").value
        Ct = Range("H2").value
        
        ratio = (Ct + Ct * 0.0291) * 100 / (10000 - Pt - Gt)
        
        Range("H5") = ratio
        
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

