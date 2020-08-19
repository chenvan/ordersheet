Attribute VB_Name = "AutoBackup"
Sub AutoBackup()
    Dim fso As Object
    Dim backupPath As String
    Dim tipsBasePath As String
    
    backupPath = "C:\备份"
    tipsBasepaht = Util.getBasePathOfTips()
    
    If Dir(backupPath, vbDirectory) = vbNullString Then
        MsgBox backupPath & " 路径不存在"
        Exit Sub
    End If
    
    ThisWorkbook.SaveCopyAs backupPath & "\备份_" & ThisWorkbook.Name
    
    Set fso = vab.CreateObject("Scripting.FileSystemObject")
    
    For Each fName In Array("defaultTips.json", "tobaccoTips.json", "cabinetTips.json")
        Call fso.CopyFile(tipsBasePath & "\" & fName, backupPath & "\" & ThisWorkbook.Name & "_" & fName, True)
    Next
    
End Sub


