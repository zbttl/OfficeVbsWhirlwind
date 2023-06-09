' FROM https://github.com/cxgreat2014/VBScript_DOC2PDF/blob/master/DOC2PDF.vbs
Dim fso,fld,Path
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
old_string=array("2021年","2021","10月31日","11月30日","12月31日","10月","11月","12月","-10-","-11-","-12-","/10/","/11/","/12/","十月","十一月","十二月","2022-2022")
new_string=array("2022年","2022","1月31日","2月28日","3月31日","1月","2月","3月","-1-","-2-","-3-","/1/","/2/","/3/","一月","二月","三月","2021-2022")
Path = fso.GetParentFolderName(WScript.ScriptFullName) '获取脚本所在文件夹字符串
Set fld=fso.GetFolder(Path) '通过路径字符串获取文件夹对象

Dim Sum,Sum2,IsChooseDelete,ThisTime
Sum = 0
Sum2 = 0
Dim LogFile
Set LogFile= fso.opentextFile("log.txt",8,true)

Dim List
Set List= fso.opentextFile("ConvertFileListDoc.txt",2,true)
Set List2= fso.opentextFile("ConvertFileListExcel.txt",2,true)

Call LogOut("开始遍历文件")
Call TreatSubFolder(fld) '调用该过程进行递归遍历该文件夹对象下的所有文件对象及子文件夹对象

Sub LogOut(msg)
    ThisTime=Now
    LogFile.WriteLine(year(ThisTime) & "-" & Month(ThisTime) & "-" & day(ThisTime) & " " & Hour(ThisTime) & ":" & Minute(ThisTime) & ":" & Second(ThisTime) & ": " & msg)
End Sub

Sub TreatSubFolder(fld) 
    Dim File
    Dim ts
    For Each File In fld.Files '遍历该文件夹对象下的所有文件对象
        IF InStr(File,"change")=0 and Mid(fso.GetFileName (File),1,2) <> "~$" Then '排除隐藏文件和已转换文件
            If UCase(fso.GetExtensionName(File)) ="DOC" or UCase(fso.GetExtensionName(File)) ="DOCX" and InStr(File,"change")=0 Then
                List.WriteLine(File.Path)
                Sum = Sum + 1
            End If
            If UCase(fso.GetExtensionName(File)) ="XLS" or UCase(fso.GetExtensionName(File)) ="XLSX" and InStr(File,"change")=0 Then
                List2.WriteLine(File.Path)
                Sum2 = Sum2 + 1
            End If
        End If
    Next
    Dim subfld
    For Each subfld In fld.SubFolders '递归遍历子文件夹对象
        TreatSubFolder subfld
    Next
End Sub
List.close
List2.close

Call LogOut("文件遍历已完成，已找到" & Sum & "个word文档," & Sum2 & "个excel文档")

If MsgBox("文件遍历已完成，已找到" & Sum & "个word文档，详细列表在" & vbCrlf & fso.GetFolder(Path).Path & "\ConvertFileListDoc.txt" & vbCrlf & "文件遍历已完成，已找到" & Sum2 & "个Excel文档，详细列表在" & vbCrlf & fso.GetFolder(Path).Path & "\ConvertFileListExcel.txt" & vbCrlf &"您可以自行修改列表以增删要转换的文档" & vbCrlf & vbCrlf & "是否将这些文档转换为PDF格式？", vbYesNo + vbInformation, "文档遍历完成") = vbYes Then
    ' If MsgBox("是否在转换完毕后删除DOC文档?", vbYesNo+vbInformation, "是否在转换完毕后删除源文档?") = vbYes Then
    '     IsChooseDelete = MsgBox("请再次确认，是否在转换完毕后删除DOC文档?", vbYesNo + vbExclamation, "是否在转换完毕后删除源文档?")
    ' End If
else
    Msgbox("已取消转换操作")
    Wscript.Quit
End If
' MsgBox "请在开始转换前退出所有Word文档避免文档占用错误发生", vbOKOnly + vbExclamation, "警告"

'创建Word对象，兼容WPS
Const wdFormatPDF = 17
On Error Resume Next
Set WordApp = CreateObject("Word.Application")
' try to connect to wps
If WordApp Is Nothing Then '兼容WPS
    Set WordApp = CreateObject("WPS.Application")
    If WordApp Is Nothing Then
        Set WordApp = CreateObject("KWPS.Application")
        If WordApp Is Nothing Then
            MsgBox "本程序依赖office 2010及以上版本，兼容WPS，" & vbCrlf & "请在使用本程序前安装office word 或WPS,否则本程序无法使用", vbCritical + vbOKOnly, "无法转换"
            WScript.Quit
        End If
    End If
End If
On Error Goto 0

WordApp.Visible=false '设置视图不可见

Sum = 0
Dim FilePath,FileLine
Set List= fso.opentextFile("ConvertFileListDoc.txt",1,true)
Do While List.AtEndOfLine <> True 
    FileLine=List.ReadLine
    If FileLine <> "" and Mid(fso.GetFileName (FilePath),1,2) <> "~$" Then
        Sum = Sum + 1 '获取用户修改后的文件列表行数
    End If
loop
List.close
' MsgBox "现在开始转换，若是在运行过程中弹出Word窗口"&vbCrlf&"请直接最小化Word窗口，不要关闭!"&vbCrlf&"请直接最小化Word窗口，不要关闭!"&vbCrlf&"请直接最小化Word窗口，不要关闭!"&vbCrlf&"重要的事情说三遍！关闭会导致脚本退出", vbOKOnly + vbExclamation, "警告"
Dim Finished
Finished = 0
Set List= fso.opentextFile("ConvertFileListDoc.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath=List.ReadLine
    If Mid(fso.GetFileName (FilePath),1,2) <> "~$" Then '不处理word临时文件
        changedocPath = fso.GetParentFolderName(FilePath) & "\" & "change" & fso.GetFileName (FilePath) 
        Set objDoc = WordApp.Documents.Open(FilePath)
        Set objSelection = WordApp.Selection
        'WordApp.Visible=false '设置视图不可见（避免运行时因为各种问题导致的可见）
        '上面这行有问题，现在遇到大批量有啥宏定义的运行起来就是一闪一闪的，还不如没有
        If WordApp.Visible = true Then
            WordApp.ActiveDocument.ActiveWindow.WindowState = 2 'wdWindowStateMinimize
        End If
        for change_string=0 To UBound(old_string)
            '参考文章https://docs.microsoft.com/en-us/office/vba/api/word.find.execute
            With objSelection.Find 
            .ClearFormatting 
            .Text = old_string(change_string)
            .Replacement.ClearFormatting 
            .Replacement.Text = new_string(change_string)
            .Forward=True
            'vbs不能用wdFindContinue，只能用数字1，否则word中如果有分页符，分页符前内容无法被替换
            .Wrap=1
            .MatchWildcards=True
            'vbs无法使用命名参数https://stackoverflow.com/questions/57463710/using-find-method-in-vbscript-for-search-word-file-can-find-character-but-not-r，replace参数只能用该种命令执行。为什么前面有10个逗号？因为在参考文章中，replace是第十一个参数
            .Execute ,,,,,,,,,,2
            End With
        Next
        objDoc.SaveAs changedocPath
        LogOut("文档" & FilePath & "已转换完成。(" & Finished & "/" & Sum & ")")
        WordApp.ActiveDocument.Close  
        Finished = Finished + 1
    End If
    ' If IsChooseDelete = vbYes Then
    '     fso.deleteFile FilePath
    '     LogOut("文件" & FilePath & "已被成功删除")
    ' End If
loop
'扫尾处理开始
List.close
' LogOut("文档转换已完成")
' LogFile.close 
'ConvertFileListDoc.txt和log.txt要自动删除的请去掉下面两行开头单引号
'fso.deleteFile "ConvertFileListDoc.txt"
'fso.deleteFile "log.txt"

' Dim Msg
' Msg = "已成功转换" & Finished & "个文件"
' ' If IsChooseDelete = vbYes Then
' '     Msg=Msg + "并成功删除源文件"
' ' End If
' MsgBox Msg & vbCrlf & "日志文件在" & fso.GetFolder(Path).Path & "\log.txt"
' ' Set fso = nothing
WordApp.Quit


' If MsgBox("文件遍历已完成，已找到" & Sum2 & "个Excel文档，详细列表在" & vbCrlf & fso.GetFolder(Path).Path & "\ConvertFileListExcel.txt" & vbCrlf & "您可以自行修改列表以增删要转换的文档" & vbCrlf & vbCrlf & "是否将这些文档转换为PDF格式？", vbYesNo + vbInformation, "文档遍历完成") = vbYes Then
'     ' If MsgBox("是否在转换完毕后删除DOC文档?", vbYesNo+vbInformation, "是否在转换完毕后删除源文档?") = vbYes Then
'     '     IsChooseDelete = MsgBox("请再次确认，是否在转换完毕后删除DOC文档?", vbYesNo + vbExclamation, "是否在转换完毕后删除源文档?")
'     ' End If
' else
'     Msgbox("已取消转换操作")
'     Wscript.Quit
' End If
' MsgBox "请在开始转换前退出所有Word文档避免文档占用错误发生", vbOKOnly + vbExclamation, "警告"

'创建Word对象，兼容WPS
On Error Resume Next
Set ExcelApp = CreateObject("Excel.Application")
' try to connect to wps
If ExcelApp Is Nothing Then '兼容WPS
    Set ExcelApp = CreateObject("WPS.Application")
    If ExcelApp Is Nothing Then
        Set ExcelApp = CreateObject("KWPS.Application")
        If ExcelApp Is Nothing Then
            MsgBox "本程序依赖office 2010及以上版本，兼容WPS，" & vbCrlf & "请在使用本程序前安装office word 或WPS,否则本程序无法使用", vbCritical + vbOKOnly, "无法转换"
            WScript.Quit
        End If
    End If
End If
On Error Goto 0

ExcelApp.Visible=false '设置视图不可见

Sum2 = 0
Dim FilePath2,FileLine2
Set List= fso.opentextFile("ConvertFileListExcel.txt",1,true)
Do While List.AtEndOfLine <> True 
    FileLine2=List.ReadLine
    If FileLine2 <> "" and Mid(fso.GetFileName (FilePath),1,2)  <> "~$" Then
        Sum2 = Sum2 + 1 '获取用户修改后的文件列表行数
    End If
loop
List.close
' MsgBox "现在开始转换，若是在运行过程中弹出Excel窗口"&vbCrlf&"请直接最小化Excel窗口，不要关闭!"&vbCrlf&"请直接最小化Excel窗口，不要关闭!"&vbCrlf&"请直接最小化Excel窗口，不要关闭!"&vbCrlf&"重要的事情说三遍！关闭会导致脚本退出", vbOKOnly + vbExclamation, "警告"
Dim Finished2
Finished2 = 0
Set List= fso.opentextFile("ConvertFileListExcel.txt",1,true)
Do While List.AtEndOfLine <> True 
    FilePath2=List.ReadLine
    If Mid(fso.GetFileName (FilePath),1,2)  <> "~$" Then '不处理Excel临时文件
        changeExcelPath = fso.GetParentFolderName(FilePath2) & "\" & "change" & fso.GetFileName (FilePath2) 
        Set objxls = ExcelApp.Workbooks.Open(FilePath2)
        'ExcelApp.Visible=false '设置视图不可见（避免运行时因为各种问题导致的可见）
        '上面这行有问题，现在遇到大批量有啥宏定义的运行起来就是一闪一闪的，还不如没有
        If ExcelApp.Visible = true Then
            ExcelApp.ActiveWorkbook.ActiveWindow.WindowState = 2 'wdWindowStateMinimize
        End If
        objcount = objxls.Sheets.Count
        for objcount_index=1 To objcount
            Set objWorksheet = objxls.Sheets(objcount_index)
            ' 循环替换字符串
            columns_count = objWorksheet.UsedRange.Columns.Count
            row_count = objWorksheet.UsedRange.Rows.Count
            for change_string=0 To UBound(old_string)
                objWorksheet.Range(objWorksheet.cells(1,1),objWorksheet.cells(row_count,columns_count)).Replace old_string(change_string), new_string(change_string)
            Next
        Next
        objxls.SaveAs changeExcelPath
        LogOut("文档" & FilePath2 & "已转换完成。(" & Finished2 & "/" & Sum2 & ")")
        ExcelApp.ActiveWorkbook.Close  
        Finished2 = Finished2 + 1
    End If
    ' If IsChooseDelete = vbYes Then
    '     fso.deleteFile FilePath2
    '     LogOut("文件" & FilePath2 & "已被成功删除")
    ' End If
loop
'扫尾处理开始
List.close
LogOut("文档转换已完成")
LogFile.close 
'ConvertFileListExcel.txt和log.txt要自动删除的请去掉下面两行开头单引号
'fso.deleteFile "ConvertFileListExcel.txt"
'fso.deleteFile "log.txt"

Dim Msg2
Msg2 = "已成功转换" & Finished & "个Word文件" & vbCrlf &"已成功转换" & Finished2 & "个Excel文件"
' If IsChooseDelete = vbYes Then
'     Msg2=Msg2 + "并成功删除源文件"
' End If
MsgBox Msg2 & vbCrlf & "日志文件在" & fso.GetFolder(Path).Path & "\log.txt"
Set fso = nothing
ExcelApp.Quit

Wscript.Quit
