' FROM https://github.com/cxgreat2014/VBScript_DOC2PDF/blob/master/DOC2PDF.vbs
Dim fso,fld,Path
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
old_string=array("2021��","2021","10��31��","11��30��","12��31��","10��","11��","12��","-10-","-11-","-12-","/10/","/11/","/12/","ʮ��","ʮһ��","ʮ����","2022-2022")
new_string=array("2022��","2022","1��31��","2��28��","3��31��","1��","2��","3��","-1-","-2-","-3-","/1/","/2/","/3/","һ��","����","����","2021-2022")
Path = fso.GetParentFolderName(WScript.ScriptFullName) '��ȡ�ű������ļ����ַ���
Set fld=fso.GetFolder(Path) 'ͨ��·���ַ�����ȡ�ļ��ж���

Dim Sum,Sum2,IsChooseDelete,ThisTime
Sum = 0
Sum2 = 0
Dim LogFile
Set LogFile= fso.opentextFile("log.txt",8,true)

Dim List
Set List= fso.opentextFile("ConvertFileListDoc.txt",2,true)
Set List2= fso.opentextFile("ConvertFileListExcel.txt",2,true)

Call LogOut("��ʼ�����ļ�")
Call TreatSubFolder(fld) '���øù��̽��еݹ�������ļ��ж����µ������ļ��������ļ��ж���

Sub LogOut(msg)
    ThisTime=Now
    LogFile.WriteLine(year(ThisTime) & "-" & Month(ThisTime) & "-" & day(ThisTime) & " " & Hour(ThisTime) & ":" & Minute(ThisTime) & ":" & Second(ThisTime) & ": " & msg)
End Sub

Sub TreatSubFolder(fld) 
    Dim File
    Dim ts
    For Each File In fld.Files '�������ļ��ж����µ������ļ�����
        If UCase(fso.GetExtensionName(File)) ="DOC" or UCase(fso.GetExtensionName(File)) ="DOCX" Then
            List.WriteLine(File.Path)
            Sum = Sum + 1
        End If
        If UCase(fso.GetExtensionName(File)) ="XLS" or UCase(fso.GetExtensionName(File)) ="XLSX" Then
            List2.WriteLine(File.Path)
            Sum2 = Sum2 + 1
        End If
    Next
    Dim subfld
    For Each subfld In fld.SubFolders '�ݹ�������ļ��ж���
        TreatSubFolder subfld
    Next
End Sub
List.close
List2.close

Call LogOut("�ļ���������ɣ����ҵ�" & Sum & "��word�ĵ�," & Sum2 & "��excel�ĵ�")

If MsgBox("�ļ���������ɣ����ҵ�" & Sum & "��word�ĵ�����ϸ�б���" & vbCrlf & fso.GetFolder(Path).Path & "\ConvertFileListDoc.txt" & vbCrlf & "�ļ���������ɣ����ҵ�" & Sum2 & "��Excel�ĵ�����ϸ�б���" & vbCrlf & fso.GetFolder(Path).Path & "\ConvertFileListExcel.txt" & vbCrlf &"�����������޸��б�����ɾҪת�����ĵ�" & vbCrlf & vbCrlf & "�Ƿ���Щ�ĵ�ת��ΪPDF��ʽ��", vbYesNo + vbInformation, "�ĵ��������") = vbYes Then
    ' If MsgBox("�Ƿ���ת����Ϻ�ɾ��DOC�ĵ�?", vbYesNo+vbInformation, "�Ƿ���ת����Ϻ�ɾ��Դ�ĵ�?") = vbYes Then
    '     IsChooseDelete = MsgBox("���ٴ�ȷ�ϣ��Ƿ���ת����Ϻ�ɾ��DOC�ĵ�?", vbYesNo + vbExclamation, "�Ƿ���ת����Ϻ�ɾ��Դ�ĵ�?")
    ' End If
else
    Msgbox("��ȡ��ת������")
    Wscript.Quit
End If
' MsgBox "���ڿ�ʼת��ǰ�˳�����Word�ĵ������ĵ�ռ�ô�����", vbOKOnly + vbExclamation, "����"

'����Word���󣬼���WPS
Const wdFormatPDF = 17
On Error Resume Next
Set WordApp = CreateObject("Word.Application")
' try to connect to wps
If WordApp Is Nothing Then '����WPS
    Set WordApp = CreateObject("WPS.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("KWPS.Application")
        If WordApp Is Nothing Then
            MsgBox "����������office 2010�����ϰ汾������WPS��" & vbCrlf & "����ʹ�ñ�����ǰ��װoffice word ��WPS,���򱾳����޷�ʹ��", vbCritical + vbOKOnly, "�޷�ת��"
            WScript.Quit
        End If
    End If
End If
On Error Goto 0

WordApp.Visible=false '������ͼ���ɼ�

Sum = 0
Dim FilePath,FileLine
Set List= fso.opentextFile("ConvertFileListDoc.txt",1,true)
Do While List.AtEndOfLine <> True 
    FileLine=List.ReadLine
    If FileLine <> "" and Mid(fso.GetFileName (FilePath),1,2) <> "~$" Then
        Sum = Sum + 1 '��ȡ�û��޸ĺ���ļ��б�����
    End If
loop
List.close
' MsgBox "���ڿ�ʼת�������������й����е���Word����"&vbCrlf&"��ֱ����С��Word���ڣ���Ҫ�ر�!"&vbCrlf&"��ֱ����С��Word���ڣ���Ҫ�ر�!"&vbCrlf&"��ֱ����С��Word���ڣ���Ҫ�ر�!"&vbCrlf&"��Ҫ������˵���飡�رջᵼ�½ű��˳�", vbOKOnly + vbExclamation, "����"
Dim Finished
Finished = 0
Set List= fso.opentextFile("ConvertFileListDoc.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath=List.ReadLine
    If Mid(fso.GetFileName (FilePath),1,2) <> "~$" Then '������word��ʱ�ļ�
        changedocPath = fso.GetParentFolderName(FilePath) & "\" & "change" & fso.GetFileName (FilePath) 
        Set objDoc = WordApp.Documents.Open(FilePath)
        Set objSelection = WordApp.Selection
        'WordApp.Visible=false '������ͼ���ɼ�����������ʱ��Ϊ�������⵼�µĿɼ���
        '�������������⣬����������������ɶ�궨���������������һ��һ���ģ�������û��
        If WordApp.Visible = true Then
            WordApp.ActiveDocument.ActiveWindow.WindowState = 2 'wdWindowStateMinimize
        End If
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
        objDoc.SaveAs changedocPath
        LogOut("�ĵ�" & FilePath & "��ת����ɡ�(" & Finished & "/" & Sum & ")")
        WordApp.ActiveDocument.Close  
        Finished = Finished + 1
    End If
    ' If IsChooseDelete = vbYes Then
    '     fso.deleteFile FilePath
    '     LogOut("�ļ�" & FilePath & "�ѱ��ɹ�ɾ��")
    ' End If
loop
'ɨβ����ʼ
List.close
' LogOut("�ĵ�ת�������")
' LogFile.close 
'ConvertFileListDoc.txt��log.txtҪ�Զ�ɾ������ȥ���������п�ͷ������
'fso.deleteFile "ConvertFileListDoc.txt"
'fso.deleteFile "log.txt"

' Dim Msg
' Msg = "�ѳɹ�ת��" & Finished & "���ļ�"
' ' If IsChooseDelete = vbYes Then
' '     Msg=Msg + "���ɹ�ɾ��Դ�ļ�"
' ' End If
' MsgBox Msg & vbCrlf & "��־�ļ���" & fso.GetFolder(Path).Path & "\log.txt"
' ' Set fso = nothing
WordApp.Quit


' If MsgBox("�ļ���������ɣ����ҵ�" & Sum2 & "��Excel�ĵ�����ϸ�б���" & vbCrlf & fso.GetFolder(Path).Path & "\ConvertFileListExcel.txt" & vbCrlf & "�����������޸��б�����ɾҪת�����ĵ�" & vbCrlf & vbCrlf & "�Ƿ���Щ�ĵ�ת��ΪPDF��ʽ��", vbYesNo + vbInformation, "�ĵ��������") = vbYes Then
'     ' If MsgBox("�Ƿ���ת����Ϻ�ɾ��DOC�ĵ�?", vbYesNo+vbInformation, "�Ƿ���ת����Ϻ�ɾ��Դ�ĵ�?") = vbYes Then
'     '     IsChooseDelete = MsgBox("���ٴ�ȷ�ϣ��Ƿ���ת����Ϻ�ɾ��DOC�ĵ�?", vbYesNo + vbExclamation, "�Ƿ���ת����Ϻ�ɾ��Դ�ĵ�?")
'     ' End If
' else
'     Msgbox("��ȡ��ת������")
'     Wscript.Quit
' End If
' MsgBox "���ڿ�ʼת��ǰ�˳�����Word�ĵ������ĵ�ռ�ô�����", vbOKOnly + vbExclamation, "����"

'����Word���󣬼���WPS
On Error Resume Next
Set ExcelApp = CreateObject("Excel.Application")
' try to connect to wps
If ExcelApp Is Nothing Then '����WPS
    Set ExcelApp = CreateObject("WPS.Application")
    If ExcelApp Is Nothing Then
        Set ExcelApp = CreateObject("KWPS.Application")
        If ExcelApp Is Nothing Then
            MsgBox "����������office 2010�����ϰ汾������WPS��" & vbCrlf & "����ʹ�ñ�����ǰ��װoffice word ��WPS,���򱾳����޷�ʹ��", vbCritical + vbOKOnly, "�޷�ת��"
            WScript.Quit
        End If
    End If
End If
On Error Goto 0

ExcelApp.Visible=false '������ͼ���ɼ�

Sum2 = 0
Dim FilePath2,FileLine2
Set List= fso.opentextFile("ConvertFileListExcel.txt",1,true)
Do While List.AtEndOfLine <> True 
    FileLine2=List.ReadLine
    If FileLine2 <> "" and Mid(fso.GetFileName (FilePath),1,2)  <> "~$" Then
        Sum2 = Sum2 + 1 '��ȡ�û��޸ĺ���ļ��б�����
    End If
loop
List.close
' MsgBox "���ڿ�ʼת�������������й����е���Excel����"&vbCrlf&"��ֱ����С��Excel���ڣ���Ҫ�ر�!"&vbCrlf&"��ֱ����С��Excel���ڣ���Ҫ�ر�!"&vbCrlf&"��ֱ����С��Excel���ڣ���Ҫ�ر�!"&vbCrlf&"��Ҫ������˵���飡�رջᵼ�½ű��˳�", vbOKOnly + vbExclamation, "����"
Dim Finished2
Finished2 = 0
Set List= fso.opentextFile("ConvertFileListExcel.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath2=List.ReadLine
    If Mid(fso.GetFileName (FilePath),1,2)  <> "~$" Then '������Excel��ʱ�ļ�
        changeExcelPath = fso.GetParentFolderName(FilePath2) & "\" & "change" & fso.GetFileName (FilePath2) 
        Set objxls = ExcelApp.Workbooks.Open(FilePath2)
        'ExcelApp.Visible=false '������ͼ���ɼ�����������ʱ��Ϊ�������⵼�µĿɼ���
        '�������������⣬����������������ɶ�궨���������������һ��һ���ģ�������û��
        If ExcelApp.Visible = true Then
            ExcelApp.ActiveWorkbook.ActiveWindow.WindowState = 2 'wdWindowStateMinimize
        End If
        objcount = objxls.Sheets.Count
        for objcount_index=1 To objcount
            Set objWorksheet = objxls.Sheets(objcount_index)
            ' ѭ���滻�ַ���
            columns_count = objWorksheet.UsedRange.Columns.Count
            row_count = objWorksheet.UsedRange.Rows.Count
            for change_string=0 To UBound(old_string)
                objWorksheet.Range(objWorksheet.cells(1,1),objWorksheet.cells(row_count,columns_count)).Replace old_string(change_string), new_string(change_string)
            Next
        Next
        objxls.SaveAs changeExcelPath
        LogOut("�ĵ�" & FilePath2 & "��ת����ɡ�(" & Finished2 & "/" & Sum2 & ")")
        ExcelApp.ActiveWorkbook.Close  
        Finished2 = Finished2 + 1
    End If
    ' If IsChooseDelete = vbYes Then
    '     fso.deleteFile FilePath2
    '     LogOut("�ļ�" & FilePath2 & "�ѱ��ɹ�ɾ��")
    ' End If
loop
'ɨβ����ʼ
List.close
LogOut("�ĵ�ת�������")
LogFile.close 
'ConvertFileListExcel.txt��log.txtҪ�Զ�ɾ������ȥ���������п�ͷ������
'fso.deleteFile "ConvertFileListExcel.txt"
'fso.deleteFile "log.txt"

Dim Msg2
Msg2 = "�ѳɹ�ת��" & Finished & "��Word�ļ�" & vbCrlf &"�ѳɹ�ת��" & Finished2 & "��Excel�ļ�"
' If IsChooseDelete = vbYes Then
'     Msg2=Msg2 + "���ɹ�ɾ��Դ�ļ�"
' End If
MsgBox Msg2 & vbCrlf & "��־�ļ���" & fso.GetFolder(Path).Path & "\log.txt"
Set fso = nothing
ExcelApp.Quit

Wscript.Quit
