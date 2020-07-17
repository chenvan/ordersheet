Private Sub Worksheet_Change(ByVal target As Range)
    Dim timeFormatColumns As Variant
    timeFormatColumns = Array(7, 8)
    
    If Not IsArray(target.value) Then
        Dim beginTimeColumn, endTimeColumn As Integer
        Dim currentTobacco, nextTobacco As String
        
        beginTimeColumn = 7
        endTimeColumn = 8
        
        If Util.IsInArray(target.Column, timeFormatColumns) And Util.IsNum(target.value) Then
            'inputTime's column is time format, which is a double type
            'we input an int, but it will change to double type
            
            If Util.IsIntType(target.value) Then
                Application.EnableEvents = False
                target.value = Others.ParseInputTime(target.value)
                Application.EnableEvents = True
            End If
            
        End If
        
        If target.Column = beginTimeColumn And Util.IsNum(target.value) Then
            '��ʼʱ���д�������ʼ��������
            
            currentTobacco = target.Offset(0, 3 - target.Column).value
            
            'Debug.Print (currentTobacco)
            
            Util.shedule currentTobacco, target.value, 2, 0
        
        End If
        
        If target.Column = endTimeColumn And Util.IsNum(target.value) Then
            '����ʱ���д�����������������, ��һ����ʼǰ��������
            
            'ע����һ�����ǿհ׵����
            
            currentTobacco = target.Offset(0, 3 - target.Column).value
            nextTobacco = target.Offset(1, 3 - target.Column).value
            
            If nextTobacco = "" Then
                '����д��һ���ε�����,��������д��һ���εĽ���ʱ��
                'Debug.Print "�հ��ƺ�"
                Util.speakAsync "������һ�����ƺ�, ��������д�����س�����ʱ��"
                Exit Sub
            End If
            
            '����������
            Util.shedule currentTobacco, target.value, 3, 0
            
            '��һ����ʼǰ����
            Util.shedule nextTobacco, target.value, 1, 0
        End If
        
    End If
    
End Sub