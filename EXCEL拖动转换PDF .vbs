'Convert .xls or .xlsx to .pdf files via Send To menu
Set fso = CreateObject("Scripting.FileSystemObject")
For i= 0 To WScript.Arguments.Count -1
   ' ��ȡ�����doc/excel������office���°�ɰ�wps
   xlsPath = WScript.Arguments(i)
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
    fso.GetBaseName(xlspath) & "change.pdf"
      objExcel.Visible = False
      Set objxls = objExcel.Workbooks.open(xlsPath)
      ' ת��Ϊ pdf
      objExcel.ActiveSheet.ExportAsFixedFormat 0, xlsPath ,0, 1, 0,,,0
      objExcel.ActiveWorkbook.Close
      objExcel.Application.Quit
   End If   
Next