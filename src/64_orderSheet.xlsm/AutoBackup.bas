Attribute VB_Name = "AutoBackup"
Sub AutoBackup()
    Dim backupPath As String
    backupPath = "C:\����"
    
    If Dir(backupPath, vbDirectory) = vbNullString Then
        MsgBox backupPath & " ·��������"
        Exit Sub
    End If
    
    ThisWorkbook.SaveCopyAs backupPath & "\����_" & ThisWorkbook.Name
End Sub
