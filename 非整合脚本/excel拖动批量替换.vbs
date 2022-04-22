'excel拖动批量替换
Set fso = CreateObject("Scripting.FileSystemObject")
'要替换的字符串们
old_string=array("2021","10月31日","11月30日","12月31日","10月","11月","12月","-10-","-11-","-12-","/10/","/11/","/12/")
new_string=array("2022","1月31日","2月28日","3月31日","1月","2月","3月","-1-","-2-","-3-","/1/","/2/","/3/")

For path_index= 0 To WScript.Arguments.Count -1
    ' 读取拖入的doc/excel，兼容office和新版旧版wps
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
        ' 处理完后的文件文件名
        changexlsPath = fso.GetParentFolderName(xlsPath) & "\" & "change" & fso.GetFileName (xlspath) 
        
        objExcel.Visible = False
        Set objxls = objExcel.Workbooks.open(xlsPath)
        '遍历所有工作簿
        objcount = objxls.Sheets.Count
        for objcount_index=1 To objcount
            Set objWorksheet = objxls.Sheets(objcount_index)
            ' 循环替换字符串
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