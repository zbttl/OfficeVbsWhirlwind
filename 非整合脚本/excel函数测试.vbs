'excel��������
Set fso = CreateObject("Scripting.FileSystemObject")


For path_index= 0 To WScript.Arguments.Count -1
    ' ��ȡ�����doc/excel������office���°�ɰ�wps
    xlsPath = WScript.Arguments(path_index)
    xlsPath = fso.GetAbsolutePathName(xlsPath)
    If LCase(Right(xlsPath, 4)) = ".xls" Or LCase(Right(xlsPath, 5)) = ".xlsx" Then
        Set objExcel = CreateObject("Excel.Application")
            If objExcel Is Nothing Then '����WPS
                Set objExcel = CreateObject("WPS.Application")
                If objExcel Is Nothing Then
                    Set objExcel = CreateObject("KWPS.Application")
                    If objExcel Is Nothing Then
                            MsgBox "����������office 2010�����ϰ汾������WPS��" & vbCrlf & "����ʹ�ñ�����ǰ��װoffice word ��WPS,���򱾳����޷�ʹ��", vbCritical + vbOKOnly, "�޷�ת��"
                            WScript.Quit
                    End If
                End If
            End If
        ' ���������ļ��ļ���
        changexlsPath = fso.GetParentFolderName(xlsPath) & "\" & "change" & fso.GetFileName (xlspath) 
        
        objExcel.Visible = False
        Set objxls = objExcel.Workbooks.open(xlsPath)
        
        objcount = objxls.Sheets.Count
        for objcount_index=1 To objcount
            Set objWorksheet = objxls.Sheets(objcount_index)
            ' ѭ���滻�ַ���
            MsgBox objWorksheet.Name
        Next

        objxls.Close
        objExcel.Quit
    End If   
Next