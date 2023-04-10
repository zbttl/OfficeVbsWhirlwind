'Convert .xls or .xlsx to .pdf files via Send To menu
'replace vba's StrConv function to vb's StrConv function
function find1(findstr)
    set rep1=new regexp
    rep1.Global=true
    rep1.IgnoreCase=true
    rep1.Pattern="[\u4E00-\u9FA5]"
    set str1=rep1.Execute(findstr)
    for each i in str1
        lens=lens+1
    next
    find1=lens + len(findstr)
end function

Function CountLines(strInput)
    Dim originalLength, noLineBreaksLength
    originalLength = Len(strInput)
    noLineBreaksLength = Len(Replace(strInput, Chr(10), ""))
    CountLines = originalLength - noLineBreaksLength + 1
End Function

Set fso = CreateObject("Scripting.FileSystemObject")
For path_index= 0 To WScript.Arguments.Count -1
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
        changexlsPath = fso.GetParentFolderName(xlsPath) & "\" & "change" & fso.GetFileName (xlspath) 
        objExcel.Visible = False
        Set objxls = objExcel.Workbooks.open(xlsPath)
        objcount = objxls.Sheets.Count
        for objcount_index=1 To objcount
            Set objWorksheet = objxls.Sheets(objcount_index)
            '��¼ÿ�еĿ��
            columns_count = objWorksheet.UsedRange.Columns.Count
            ReDim columns_width(columns_count-1)
            For columns_count_index=0 To columns_count-1
            If (objWorksheet.Columns(columns_count_index + 1).ColumnWidth) Mod 2 = 0 Then
            columns_width(columns_count_index) = objWorksheet.Columns(columns_count_index + 1).ColumnWidth
            Else
            columns_width(columns_count_index) = objWorksheet.Columns(columns_count_index + 1).ColumnWidth - 1
            End If
            Next
            '����ÿ�е����߶�
            dim Columns_line,Columns_line_Max
            For row_index = 3 To objWorksheet.UsedRange.Rows.Count
            Columns_line_Max = 0
            For columns_index = 0 To columns_count - 1
            'CInt����ȡ��https://blog.csdn.net/iamlaosong/article/details/49333779
            'VBA��lenB:https://blog.csdn.net/iamlaosong/article/details/49333779
            Columns_line=CInt(find1(objWorksheet.Cells(row_index,columns_index+1))/columns_width(columns_index)+0.5)
            If Columns_line>Columns_line_Max Then
            Columns_line_Max = Columns_line
            End If
            Next
            On Error Resume Next ' ���ô�����
            objWorksheet.Rows(row_index).RowHeight = Columns_line_Max * 15+15
            '��������и߻ᱨ��
            If Err.Number <> 0 Then
                'WScript.Echo "�У�"&row_index&"�У�"&columns_index+1&"RowHeight:"&objWorksheet.Rows(row_index).RowHeight&"Columns_line_Max"&Columns_line_Max&"��������" & Err.Description & " (������룺" & Err.Number & ")"
                objWorksheet.Rows(row_index).RowHeight = 409
                Err.Clear ' �����������Ӱ���������
            End If
            Next
        Next 
        objxls.saveas changexlsPath
        objxls.Close
        objExcel.Quit
    End If   
Next