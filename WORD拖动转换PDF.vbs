'Convert .doc or .docx to .pdf files via Send To menu
Set fso = CreateObject("Scripting.FileSystemObject")
For i= 0 To WScript.Arguments.Count -1
   ' 读取doc/excel，兼容office和新版旧版wps
   docPath = WScript.Arguments(i)
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
Next