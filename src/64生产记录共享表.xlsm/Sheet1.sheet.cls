Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(8, 9)
    
    If Not IsArray(target.value) Then
    
        If Util.IsInArray(target.Column, timeFormatColumns) And Util.IsNum(target.value) Then
            'inputTime's column is time format, which is a double type
            'we input an int, but it will change to double type
            If Util.IsIntType(target.value) Then
                target.value = Others.ParseInputTime(target.value)
            End If
            
        End If
    
    End If
    
End Sub