Attribute VB_Name = "Others"
Function ParseInputTime(ByVal abbrTime As Integer) As Date
    
    Dim hour As Integer
    Dim min As Integer
    
    If inputTime < 2400 Then
        hour = Int(abbrTime / 100)
        min = Int(abbrTime Mod 100)
    
        ParseInputTime = TimeSerial(hour, min, 0)
    Else
        MsgBox "��ֵ����2400,�滻Ϊ����ʱ��"
        ParseInputTime = Now()
    End If

End Function

