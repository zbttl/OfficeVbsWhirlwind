'word�϶������滻
Set fso = CreateObject("Scripting.FileSystemObject")
'Ҫ�滻���ַ�����
old_string=array("2021","10��31��","11��30��","12��31��","10��","11��","12��")
new_string=array("2022","1��31��","2��28��","3��31��","1��","2��","3��")

For i= 0 To WScript.Arguments.Count -1
   docPath = WScript.Arguments(i)
   docPath = fso.GetAbsolutePathName(docPath)
   If LCase(Right(docPath, 4)) = ".doc" Or LCase(Right(docPath, 5)) = ".docx" Then
      Set objWord = CreateObject("Word.Application")
         If objWord Is Nothing Then '����WPS
            Set objWord = CreateObject("WPS.Application")
            If objWord Is Nothing Then
               Set objWord = CreateObject("KWPS.Application")
               If objWord Is Nothing Then
                     MsgBox "����������office 2010�����ϰ汾������WPS��" & vbCrlf & "����ʹ�ñ�����ǰ��װoffice word ��WPS,���򱾳����޷�ʹ��", vbCritical + vbOKOnly, "�޷�ת��"
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