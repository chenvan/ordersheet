Attribute VB_Name = "AutoBackup"
Sub AutoBackup()
    ThisWorkbook.SaveCopyAs "C:\����\" & "����_" & ThisWorkbook.Name
End Sub
