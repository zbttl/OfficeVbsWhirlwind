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


Set fso = CreateObject("Scripting.FileSystemObject")
For path_index= 0 To WScript.Arguments.Count -1
    xlsPath = WScript.Arguments(path_index)
    xlsPath = fso.GetAbsolutePathName(xlsPath)
    If LCase(Right(xlsPath, 4)) = ".xls" Or LCase(Right(xlsPath, 5)) = ".xlsx" Then
        Set objExcel = CreateObject("Excel.Application")
            If objExcel Is Nothing Then '兼容WPS
                Set objExcel = CreateObject("WPS.Application")
                If objExcel Is Nothing Then
                    Set objExcel = CreateObject("KWPS.Application")
                    If objExcel Is Nothing Then
                            MsgBox "本程序依赖office 2010及以上版本，兼容WPS，" & vbCrlf & "请在使用本程序前安装office word 或WPS,否则本程序无法使用", vbCritical + vbOKOnly, "无法转换"
                            WScript.Quit
                    End If
                End If
            End If
        changexlsPath = fso.GetParentFolderName(xlsPath) & "\" & "change" & fso.GetFileName (xlspath) 
        objExcel.Visible = False
        Set objxls = objExcel.Workbooks.open(xlsPath)
        Set objWorksheet = objxls.Sheets(1)
        '记录每列的宽度
        columns_count = objWorksheet.UsedRange.Columns.Count
        ReDim columns_width(columns_count-1)
        For columns_count_index=0 To columns_count-1
        If (objWorksheet.Columns(columns_count_index + 1).ColumnWidth) Mod 2 = 0 Then
        columns_width(columns_count_index) = objWorksheet.Columns(columns_count_index + 1).ColumnWidth
        Else
        columns_width(columns_count_index) = objWorksheet.Columns(columns_count_index + 1).ColumnWidth - 1
        End If
        Next
        '计算每列的最大高度
        dim Columns_line,Columns_line_Max
        For row_index = 3 To objWorksheet.UsedRange.Rows.Count
        Columns_line_Max = 0
        For columns_index = 0 To columns_count - 1
        'CInt向上取整https://blog.csdn.net/iamlaosong/article/details/49333779
        'VBA的lenB:https://blog.csdn.net/iamlaosong/article/details/49333779
        Columns_line=CInt(find1(objWorksheet.Cells(row_index,columns_index+1))/columns_width(columns_index)+0.5)
        If Columns_line>Columns_line_Max Then
        Columns_line_Max = Columns_line
        End If
        Next
        objWorksheet.Rows(row_index).RowHeight = Columns_line_Max * 15+2
        Next 
        objxls.saveas changexlsPath
        objxls.Close
        objExcel.Quit
    End If   
Next