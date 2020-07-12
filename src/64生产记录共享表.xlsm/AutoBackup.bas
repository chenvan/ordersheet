Attribute VB_Name = "AutoBackup"
Sub AutoBackup()
    ThisWorkbook.SaveCopyAs "C:\备份\" & "备份_" & ThisWorkbook.Name
End Sub
