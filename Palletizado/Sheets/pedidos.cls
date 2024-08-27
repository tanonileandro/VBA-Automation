VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 ' Código para el botón de seleccionar
Private Sub btnSeleccionar_Click()
    Dim ultimaFila As Long
    Dim hojaDespachado As Worksheet

    Set hojaDespachado = ThisWorkbook.Sheets("PEDIDOS")

    Application.EnableEvents = False

    ultimaFila = hojaDespachado.Cells(hojaDespachado.Rows.Count, "A").End(xlUp).Row

    If Application.WorksheetFunction.CountA(hojaDespachado.Range("A4:A" & ultimaFila)) > 0 Then
        hojaDespachado.Rows("4:" & ultimaFila).Select
    Else
        MsgBox "No hay filas con datos para seleccionar.", vbInformation
    End If

    Application.EnableEvents = True
End Sub

' Código para el botón de borrar
Private Sub btnBorrar_Click()
    Dim ultimaFila As Long
    Dim hojaDespachado As Worksheet
    Dim seleccion As Range
    Dim respuesta As VbMsgBoxResult

    Set hojaDespachado = ThisWorkbook.Sheets("PEDIDOS")

    On Error Resume Next
    Set seleccion = Selection
    On Error GoTo 0

    If seleccion Is Nothing Then
        MsgBox "Primero debe seleccionar las filas que desea borrar.", vbExclamation
    Else
        respuesta = MsgBox("Está por borrar todos los despachos seleccionados. ¿Desea continuar?", vbYesNo + vbQuestion, "Confirmar borrado")

        If respuesta = vbYes Then
            Application.EnableEvents = False

            If Not IsEmpty(seleccion) Then
            
                ultimaFila = hojaDespachado.Cells(hojaDespachado.Rows.Count, "A").End(xlUp).Row

                seleccion.Rows.Delete
            Else
                MsgBox "No hay nada para borrar.", vbInformation
            End If

            Application.EnableEvents = True
        End If
    End If
End Sub
