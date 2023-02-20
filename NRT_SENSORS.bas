Attribute VB_Name = "NRT_SENSORS"
Option Explicit

'VERSION: 2022.03.16
'AUTHOR: PABLO GONZALEZ PILA <pablogonzalezpila@gmail.com>

'MACROS de uso temporal

Sub TEMP()

MsgBox "TEMP"

End Sub

Sub NRTZ14_CLEAR()

    Sheets("DATOS ISO").Select
    
    Range("E10:E12").Select
    Selection.ClearContents
    
    Range("E17:F24").Select
    Selection.ClearContents

    Range("E29:E30").Select
    Selection.ClearContents

    Range("E35:F45").Select
    Selection.ClearContents

    Range("E47:F57").Select
    Selection.ClearContents

    Range("E62:E63").Select
    Selection.ClearContents

    Range("E68:E78").Select
    Selection.ClearContents

    Range("E83").Select
    Selection.ClearContents

    Range("E88:E98").Select
    Selection.ClearContents

    Range("E103:E113").Select
    Selection.ClearContents

    Range("E118").Select
    Selection.ClearContents
    
    Sheets("TestReport").Select

    Range("E5").Select
    Selection.ClearContents

    Range("B5").Select
    Selection.ClearContents
    
    Range("B5").Select
    Selection.NumberFormat = "yyyy-mm-dd"
    ActiveCell.FormulaR1C1 = Now
    
    Sheets("DATOS ISO").Select
    ActiveWindow.SmallScroll Down:=-117
    
End Sub

Sub NRTZ14_COMPLETE()

'MsgBox "NRTZ14_COMPLETE"

    Sheets("DATOS ISO").Select
    
    Range("G29:G30").Select
    Selection.Copy
    Range("E29").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G62:G63").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E62").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G83").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E83").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G118").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E118").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G29:G30").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=18
    Range("G62:G63").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=36
    Range("G83").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=33
    Range("G118").Select
    Selection.ClearContents
    
    ActiveWindow.SmallScroll Down:=-117

End Sub

Sub NRTZ44_CLEAR()

    Sheets("DATOS ISO").Select
    
    Range("E10:E12").Select
    Selection.ClearContents
    
    Range("E17:F33").Select
    Selection.ClearContents

    Range("E38:E54").Select
    Selection.ClearContents

    Range("E59:F73").Select
    Selection.ClearContents

    Range("E78:E79").Select
    Selection.ClearContents

    Range("E84:E92").Select
    Selection.ClearContents

    Range("E97").Select
    Selection.ClearContents

    Range("E102:F108").Select
    Selection.ClearContents

    Range("E113:E114").Select
    Selection.ClearContents

    Range("E119").Select
    Selection.ClearContents
    
    Sheets("TestReport").Select

    Range("E5").Select
    Selection.ClearContents

    Range("B5").Select
    Selection.ClearContents
    
    Range("B5").Select
    Selection.NumberFormat = "yyyy-mm-dd"
    ActiveCell.FormulaR1C1 = Now
    
    Sheets("DATOS ISO").Select
    ActiveWindow.SmallScroll Down:=-117

End Sub

Sub NRTZ44_COMPLETE()

    Sheets("DATOS ISO").Select
    
    Range("G78:G79").Select
    Selection.Copy
    Range("E78").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G97").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E97").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G113:G114").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E113").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G119").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("E119").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G78:G79").Select
    Selection.ClearContents

    Range("G96:G98").Select
    Selection.ClearContents

    Range("G112:G119").Select
    Selection.ClearContents

    ActiveWindow.SmallScroll Down:=-117

End Sub

