Attribute VB_Name = "AutoBackup"
Sub AutoBackup()
    Dim backupPath As String
    backupPath = "C:\备份"
    
    If Dir(backupPath, vbDirectory) = vbNullString Then
        MsgBox backupPath & " 路径不存在"
        Exit Sub
    End If
    
    ThisWorkbook.SaveCopyAs backupPath & "\备份_" & ThisWorkbook.Name
End Sub
