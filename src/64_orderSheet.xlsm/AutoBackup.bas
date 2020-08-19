Attribute VB_Name = "AutoBackup"
Sub AutoBackup()
    Dim fso As Object
    Dim backupPath As String
    Dim tipsBasePath As String
    
    backupPath = "C:\����"
    tipsBasepaht = Util.getBasePathOfTips()
    
    If Dir(backupPath, vbDirectory) = vbNullString Then
        MsgBox backupPath & " ·��������"
        Exit Sub
    End If
    
    ThisWorkbook.SaveCopyAs backupPath & "\����_" & ThisWorkbook.Name
    
    Set fso = vab.CreateObject("Scripting.FileSystemObject")
    
    For Each fName In Array("defaultTips.json", "tobaccoTips.json", "cabinetTips.json")
        Call fso.CopyFile(tipsBasePath & "\" & fName, backupPath & "\" & ThisWorkbook.Name & "_" & fName, True)
    Next
    
End Sub


