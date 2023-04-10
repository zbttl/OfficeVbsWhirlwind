'Convert .doc or .docx to .pdf files via Send To menu
Set fso = CreateObject("Scripting.FileSystemObject")
For path_index= 0 To WScript.Arguments.Count -1
   ' ��ȡdoc������office���°�ɰ�wps
   docPath = WScript.Arguments(path_index)
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
   ' ��ȡexcel������office���°�ɰ�wps
   xlsPath = WScript.Arguments(path_index)
   xlsPath = fso.GetAbsolutePathName(xlsPath)
   If LCase(Right(xlsPath, 4)) = ".xls" Or LCase(Right(xlsPath, 5)) = ".xlsx" Then
      Set objExcel = CreateObject("Excel.Application")
         If objExcel Is Nothing Then '����WPS
            Set objExcel = CreateObject("WPS.Application")
            If objExcel Is Nothing Then
               Set objExcel = CreateObject("KWPS.Application")
               If objExcel Is Nothing Then
                     MsgBox "����������office 2010�����ϰ汾������WPS��" & vbCrlf & "����ʹ�ñ�����ǰ��װoffice word ��WPS,���򱾳����޷�ʹ��", vbCritical + vbOKOnly, "�޷�ת��"
                     WScript.Quit
               End If
            End If
         End If
      ' ���������ļ��ļ���
      changexlsPath = fso.GetParentFolderName(xlsPath) & "\" & _
    fso.GetBaseName(xlspath) & ".pdf"
      objExcel.Visible = False
      Set objxls = objExcel.Workbooks.open(xlsPath)
      ' ���ù�����Ϊ���򲼾ֲ���������������Ӧһҳ
      With objExcel.ActiveSheet.PageSetup
      .Orientation = 2 ' xlLandscape
      .Zoom = False
      .FitToPagesWide = 1
      .FitToPagesTall = False
      End With
      ' ת��Ϊ pdf
      objExcel.ActiveSheet.ExportAsFixedFormat 0, changexlsPath ,0, 1, 0,,,0
      
      objExcel.DisplayAlerts = False
      objExcel.ActiveWorkbook.Close
      objExcel.DisplayAlerts = True
      objExcel.Application.Quit
   End If   
Next