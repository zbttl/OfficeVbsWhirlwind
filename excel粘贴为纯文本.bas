Sub 粘贴为纯文本()
'
' 粘贴为纯文本 Macro
'
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub