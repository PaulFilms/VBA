Attribute VB_Name = "GENERAL"
Option Explicit

'VERSION: 2022.07.05
'AUTHOR: PABLO GONZALEZ PILA <pablogonzalezpila@gmail.com>

Sub CREATE_SHORTCUTS()
' https://docs.microsoft.com/en-us/office/vba/api/excel.application.onkey

    ' Asignamos el atajo [CTRL+ñ] a la función SPECIAL_PASTE
    Application.OnKey "^{ñ}", "SPECIAL_PASTE"
'    Debug.Print "CREATE_SHORTCUTS"

End Sub

Sub SPECIAL_PASTE()
' Función para realizar el "Pegado Especial" con la configuración de solo datos, sin formato

    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'    Debug.Print "SPECIAL_PASTE"
    
End Sub



