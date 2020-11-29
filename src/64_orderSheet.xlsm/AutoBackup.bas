Attribute VB_Name = "AutoBackup"
Option Explicit

Sub AutoBackup()
    Dim fso As Object
    Dim backupPath, tipsPath As String
    
    backupPath = Sheets("�趨").Range("A:A").Find("����·��").offset(0, 1).value
    tipsPath = Sheets("�趨").Range("A:A").Find("�����ļ�·��").offset(0, 1).value
    
    If Dir(backupPath, vbDirectory) = vbNullString Then
        MsgBox backupPath & " ·��������"
        Exit Sub
    End If
    
    ThisWorkbook.SaveCopyAs backupPath & "\����_" & ThisWorkbook.Name
    
    If Dir(tipsPath, vbNormal) = vbNullString Then
        MsgBox tipsPath & "�ļ�������"
        Exit Sub
    End If
    
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Call fso.CopyFile(tipsPath, backupPath & "\����_" & ThisWorkbook.Name & "_tips.json", True)

End Sub



