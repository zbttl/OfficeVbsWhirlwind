'word拖动批量替换
Set fso = CreateObject("Scripting.FileSystemObject")
'要替换的字符串们
old_string=array("2021","10月31日","11月30日","12月31日","10月","11月","12月")
new_string=array("2022","1月31日","2月28日","3月31日","1月","2月","3月")

For i= 0 To WScript.Arguments.Count -1
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
      changedocPath = fso.GetParentFolderName(docPath) & "\" & "change" & fso.GetFileName (docPath) 

      objWord.Visible = False
      Set objDoc = objWord.documents.open(docPath)
      Set objSelection = objWord.Selection
      for change_string=0 To UBound(old_string)
         With objSelection.Find 
         .ClearFormatting 
         .Text = old_string(change_string)
         .Replacement.ClearFormatting 
         .Replacement.Text = new_string(change_string)
         .Forward=True
         .Wrap=wdFindContinue
         .MatchWildcards=True
         .Execute ,,,,,,,,,,2
         End With
      Next
      objDoc.saveas changedocPath
      objDoc.Close
      ' DocWord.Close
      objWord.Quit   
   End If   
Next