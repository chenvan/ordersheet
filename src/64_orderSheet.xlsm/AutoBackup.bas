Attribute VB_Name = "AutoBackup"
Option Explicit

Sub AutoBackup()
    Dim fso As Object
    Dim backupPath, tipsPath As String
    
    backupPath = "C:\备份"
    tipsPath = Sheets("设定").range("A:A").Find("语音文件路径").offset(0, 1).value
    
    If Dir(backupPath, vbDirectory) = vbNullString Then
        MsgBox backupPath & " 路径不存在"
        Exit Sub
    End If
    
    ThisWorkbook.SaveCopyAs backupPath & "\备份_" & ThisWorkbook.Name
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Call fso.CopyFile(tipsPath, backupPath & "\" & ThisWorkbook.Name & "_tips_backup.json", True)

End Sub


