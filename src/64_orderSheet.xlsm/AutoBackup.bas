Attribute VB_Name = "AutoBackup"
Option Explicit

Sub AutoBackup()
    Dim fso As Object
    Dim backupPath, tipsPath As String
    
    backupPath = Sheets("设定").Range("A:A").Find("备份路径").offset(0, 1).value
    tipsPath = Sheets("设定").Range("A:A").Find("语音文件路径").offset(0, 1).value
    
    If Dir(backupPath, vbDirectory) = vbNullString Then
        MsgBox backupPath & " 路径不存在"
        Exit Sub
    End If
    
    ThisWorkbook.SaveCopyAs backupPath & "\备份_" & ThisWorkbook.Name
    
    If Dir(tipsPath, vbNormal) = vbNullString Then
        MsgBox tipsPath & "文件不存在"
        Exit Sub
    End If
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Call fso.CopyFile(tipsPath, backupPath & "\备份_" & ThisWorkbook.Name & "_tips.json", True)

End Sub



