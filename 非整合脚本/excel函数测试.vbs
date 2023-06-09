'excel函数测试
Set fso = CreateObject("Scripting.FileSystemObject")


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
        
        objcount = objxls.Sheets.Count
        for objcount_index=1 To objcount
            Set objWorksheet = objxls.Sheets(objcount_index)
            ' 循环替换字符串
            MsgBox objWorksheet.Name
        Next

        objxls.Close
        objExcel.Quit
    End If   
Next