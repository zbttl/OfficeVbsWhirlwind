'Convert .doc or .docx to .pdf files via Send To menu
Set fso = CreateObject("Scripting.FileSystemObject")
For path_index= 0 To WScript.Arguments.Count -1
   ' 读取doc，兼容office和新版旧版wps
   docPath = WScript.Arguments(path_index)
   docPath = fso.GetAbsolutePathName(docPath)
   If LCase(Right(docPath, 4)) = ".doc" Or LCase(Right(docPath, 5)) = ".docx" Then
      Set objWord = CreateObject("Word.Application")
         If objWord Is Nothing Then '兼容WPS
            Set objWord = CreateObject("WPS.Application")
            If objWord Is Nothing Then
               Set objWord = CreateObject("KWPS.Application")
               If objWord Is Nothing Then
                     MsgBox "本程序依赖office 2010及以上版本，兼容WPS，" & vbCrlf & "请在使用本程序前安装office word 或WPS,否则本程序无法使用", vbCritical + vbOKOnly, "无法转换"
                     WScript.Quit
               End If
            End If
         End If
      ' 处理完后的文件文件名
      pdfPath = fso.GetParentFolderName(docPath) & "\" & _
    fso.GetBaseName(docpath) & ".pdf"
      objWord.Visible = False
      Set objDoc = objWord.documents.open(docPath)
      ' 转换为 pdf
      objDoc.saveas pdfPath, 17
      objDoc.Close
      objWord.Quit   
   End If
   ' 读取excel，兼容office和新版旧版wps
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
      changexlsPath = fso.GetParentFolderName(xlsPath) & "\" & _
    fso.GetBaseName(xlspath) & ".pdf"
      objExcel.Visible = False
      Set objxls = objExcel.Workbooks.open(xlsPath)
      ' 设置工作表为横向布局并调整所有列以适应一页
      With objExcel.ActiveSheet.PageSetup
      .Orientation = 2 ' xlLandscape
      .Zoom = False
      .FitToPagesWide = 1
      .FitToPagesTall = False
      End With
      ' 转换为 pdf
      objExcel.ActiveSheet.ExportAsFixedFormat 0, changexlsPath ,0, 1, 0,,,0
      
      objExcel.DisplayAlerts = False
      objExcel.ActiveWorkbook.Close
      objExcel.DisplayAlerts = True
      objExcel.Application.Quit
   End If   
Next