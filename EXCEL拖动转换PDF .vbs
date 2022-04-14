'Convert .xls or .xlsx to .pdf files via Send To menu
Set fso = CreateObject("Scripting.FileSystemObject")
For i= 0 To WScript.Arguments.Count -1
   ' 读取拖入的doc/excel，兼容office和新版旧版wps
   xlsPath = WScript.Arguments(i)
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
      changexlsPath = fso.GetParentFolderName(xlsPath) & "\" & _
    fso.GetBaseName(xlspath) & "change.pdf"
      objExcel.Visible = False
      Set objxls = objExcel.Workbooks.open(xlsPath)
      ' 转换为 pdf
      objExcel.ActiveSheet.ExportAsFixedFormat 0, xlsPath ,0, 1, 0,,,0
      objExcel.ActiveWorkbook.Close
      objExcel.Application.Quit
   End If   
Next