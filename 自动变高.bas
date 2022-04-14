sub getColomnWidth()
    '记录每列的宽度
    columns_count = ActiveSheet.UsedRange.Columns.Count
    ReDim columns_width(columns_count-1) As Long
    For i = 0 To columns_count - 1
    If (Columns(i + 1).ColumnWidth) Mod 2 = 0 Then
    columns_width(i) = Columns(i + 1).ColumnWidth
    Else
    columns_width(i) = Columns(i + 1).ColumnWidth - 1
    End If
    Next i
    '计算每列的最大高度
    dim Columns_line,Columns_line_Max as long
    For i = 3 To ActiveSheet.UsedRange.Rows.Count
    Columns_line_Max = 0
    For j = 0 To columns_count - 1
    'CInt向上取整https://blog.csdn.net/iamlaosong/article/details/49333779
    'VBA的lenB:https://blog.csdn.net/iamlaosong/article/details/49333779
    Columns_line=CInt(LenB(StrConv(Cells(i,j+1), vbFromUnicode))/columns_width(j)+0.5)
    If Columns_line>Columns_line_Max Then
    Columns_line_Max = Columns_line
    End If
    Next j
    Worksheets(1).Rows(i).RowHeight = Columns_line_Max * 15+2
    Next i
End  Sub