Attribute VB_Name = "Tobacco"
Option Explicit


Function ParseInputCode(inputCode As String) As String()
    Dim parsedResult(0 To 2) As String
    
     Select Case Len(inputCode)
        Case Is = 7
            parsedResult(0) = Left(inputCode, 2)
            parsedResult(1) = Mid(inputCode, 3, 2)
            parsedResult(2) = Right(inputCode, 3)
        Case Is = 5
            parsedResult(0) = Format(Now(), "yy")
            parsedResult(1) = Left(inputCode, 2)
            parsedResult(2) = Right(inputCode, 3)
        Case Is <= 3
            parsedResult(0) = Format(Now(), "yy")
            parsedResult(1) = Format(Now(), "mm")
            parsedResult(2) = Format(inputCode, "000")
    End Select
    
    ParseInputCode = parsedResult

End Function


Function GetTobaccoCode(ByVal tobaccoName As String, ByVal sheetNameWithCode As String) As String
    ' sheetName's column A is tobaccoName, column B is tobaccoCode
    Dim tobaccoCode As String
    tobaccoCode = ""
    
    If tobaccoName <> "" Then
        Dim found As range
        Set found = Sheets(sheetNameWithCode).range("A:A").Find(tobaccoName, , , xlWhole)
        tobaccoCode = found.offset(0, 1).value
    End If

    GetTobaccoCode = tobaccoCode

End Function
