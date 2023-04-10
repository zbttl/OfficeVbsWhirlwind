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

Function ReplaceMultipleSpacesWithOne(inputString)
    Dim oRegExp
    Set oRegExp = New RegExp
    
    oRegExp.Pattern = " {2,}" ' 匹配两个或更多连续空格
    oRegExp.Global = True ' 替换所有匹配项
    
    ReplaceMultipleSpacesWithOne = oRegExp.Replace(inputString, " ") ' 用一个空格替换匹配的空格
End Function

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
        objcount = objxls.Sheets.Count
        for objcount_index=1 To objcount
            Set objWorksheet = objxls.Sheets(objcount_index)
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

            strCellContent = objWorksheet.Cells(row_index,columns_index+1) 
            strCellContent = Replace(strCellContent, Chr(10), " ")
            strCellContent = Replace(strCellContent, ChrW(160), " ")
            strCellContent = ReplaceMultipleSpacesWithOne(strCellContent)
            objWorksheet.Cells(row_index,columns_index+1) = strCellContent

            Columns_line=CInt(find1(objWorksheet.Cells(row_index,columns_index+1))/columns_width(columns_index)+0.5)

            ' Columns_line2=CountLines(objWorksheet.Cells(row_index,columns_index+1))
            ' If Columns_line<Columns_line2 Then
            '     Columns_line=Columns_line2
            ' End If
            If Columns_line>Columns_line_Max Then
            Columns_line_Max = Columns_line
            End If
            Next
            On Error Resume Next ' 启用错误处理
            objWorksheet.Rows(row_index).RowHeight = Columns_line_Max * 15+15
            '超过最大行高会报错
            If Err.Number <> 0 Then
                'WScript.Echo "行："&row_index&"列："&columns_index+1&"RowHeight:"&objWorksheet.Rows(row_index).RowHeight&"Columns_line_Max"&Columns_line_Max&"发生错误：" & Err.Description & " (错误代码：" & Err.Number & ")"
                objWorksheet.Rows(row_index).RowHeight = 409
                Err.Clear ' 清除错误，以免影响后续代码
            End If
            Next
        Next 
        objxls.saveas changexlsPath
        objxls.Close
        objExcel.Quit
    End If   
Next