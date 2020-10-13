Option Explicit

Private Sub Workbook_Open()
    '自动备份
    
    Application.OnTime TimeValue("12:10:00"), "AutoBackup.AutoBackup"
    Application.OnTime TimeValue("22:10:00"), "AutoBackup.AutoBackup"
    
    'Util.loadDefaultTips
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
    Dim sheetNames As Variant
    Dim settingSheetName As String
    
    sheetNames = Array("切烘加香段", "回潮段", "加料段", "HDT段")
    
    If Not IsArray(target.value) Then
        
        If Util.IsInArray(Sh.Name, sheetNames) And target.Column = 4 And Util.IsNum(target.value) Then
            Dim tobaccoCode As String
            tobaccoCode = Tobacco.GetTobaccoCode(target.offset(0, -1).value, "设定")
            
            If tobaccoCode <> "" Then
                Dim segments() As String
                segments = Tobacco.ParseInputCode(target.value)
                
                If Not Util.IsInArray("", segments) Then
                    target.value = segments(0) & segments(1) & tobaccoCode & segments(2)
                End If
                
            End If
            
        End If
        
    End If
    
End Sub


Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal target As Range)
    'hightlight selectd cell's row and  column, need to set rule first at home tab
    If Application.CutCopyMode = False Then
        Application.Calculate
    End If
    
End Sub


