'excel�϶������滻
Set fso = CreateObject("Scripting.FileSystemObject")
'Ҫ�滻���ַ�����
old_string=array("2021","10��31��","11��30��","12��31��","10��","11��","12��","-10-","-11-","-12-","/10/","/11/","/12/")
new_string=array("2022","1��31��","2��28��","3��31��","1��","2��","3��","-1-","-2-","-3-","/1/","/2/","/3/")

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
        '�������й�����
        objcount = objxls.Sheets.Count
        for objcount_index=1 To objcount
            Set objWorksheet = objxls.Sheets(objcount_index)
            ' ѭ���滻�ַ���
            columns_count = objWorksheet.UsedRange.Columns.Count
            row_count = objWorksheet.UsedRange.Rows.Count
            for change_string=0 To UBound(old_string)
                objWorksheet.Range(objWorksheet.cells(1,1),objWorksheet.cells(row_count,columns_count)).Replace old_string(change_string), new_string(change_string)
            Next
        Next
        objxls.saveas changexlsPath
        objxls.Close
        objExcel.Quit
    End If   
Next