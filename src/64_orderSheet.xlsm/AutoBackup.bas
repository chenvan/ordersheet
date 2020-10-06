Attribute VB_Name = "AutoBackup"
Option Explicit

Sub AutoBackup()
    Dim fso As Object
    Dim backupPath, tipsPath As String
    
    backupPath = "C:\����"
    tipsPath = Sheets("�趨").range("A:A").Find("�����ļ�·��").offset(0, 1).value
    
    If Dir(backupPath, vbDirectory) = vbNullString Then
        MsgBox backupPath & " ·��������"
        Exit Sub
    End If
    
    ThisWorkbook.SaveCopyAs backupPath & "\����_" & ThisWorkbook.Name
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Call fso.CopyFile(tipsPath, backupPath & "\" & ThisWorkbook.Name & "_tips_backup.json", True)

End Sub


