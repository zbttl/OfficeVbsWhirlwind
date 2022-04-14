'Convert .doc or .docx to .pdf files via Send To menu
Set fso = CreateObject("Scripting.FileSystemObject")
For i= 0 To WScript.Arguments.Count -1
   ' ��ȡdoc/excel������office���°�ɰ�wps
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
      ' ���������ļ��ļ���
      pdfPath = fso.GetParentFolderName(docPath) & "\" & _
    fso.GetBaseName(docpath) & ".pdf"
      objWord.Visible = False
      Set objDoc = objWord.documents.open(docPath)
      ' ת��Ϊ pdf
      objDoc.saveas pdfPath, 17
      objDoc.Close
      objWord.Quit   
   End If   
Next