VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim hojaDespachado As Worksheet
    Dim hojaPalletizador As Worksheet
    Dim codigoBarra As String
    Dim buscarRango As Range
    Dim coincidencia As Range
    
    Set hojaDespachado = ThisWorkbook.Sheets("DESPACHADO")
    Set hojaPalletizador = ThisWorkbook.Sheets("PALLETIZADO")

    If Target.Cells.Count = 1 And Not IsEmpty(Target.Value) Then

        If Not Intersect(Target, hojaDespachado.Range("A:A")) Is Nothing Then
            codigoBarra = Target.Value

            Set buscarRango = hojaPalletizador.Range("A:A")

            Set coincidencia = buscarRango.Find(What:=codigoBarra, LookIn:=xlValues, LookAt:=xlWhole)

            If Not coincidencia Is Nothing Then
                Application.EnableEvents = False

                hojaPalletizador.Rows(coincidencia.Row).Copy hojaDespachado.Range("A" & Target.Row)


                hojaDespachado.Range("K" & Target.Row).Value = "Despachado"
                hojaDespachado.Range("K" & Target.Row).Interior.Color = RGB(0, 255, 0)

                hojaDespachado.Range("B" & Target.Row).Value = Now

                hojaPalletizador.Rows(coincidencia.Row).Delete

                Application.EnableEvents = True
            Else
        
                hojaDespachado.Rows(Target.Row).ClearContents
            End If
        End If
    End If
End Sub

 ' BOTON PARA SELECCIONAR TODO
 
Private Sub btnSeleccionar_Click()
    Dim ultimaFila As Long
    Dim hojaDespachado As Worksheet

    Set hojaDespachado = ThisWorkbook.Sheets("DESPACHADO")

    Application.EnableEvents = False

    ultimaFila = hojaDespachado.Cells(hojaDespachado.Rows.Count, "A").End(xlUp).Row

    If Application.WorksheetFunction.CountA(hojaDespachado.Range("A4:A" & ultimaFila)) > 0 Then

        hojaDespachado.Rows("4:" & ultimaFila).Select
    Else
        MsgBox "No hay filas con datos para seleccionar.", vbInformation
    End If

    Application.EnableEvents = True
End Sub

' BOTON PARA BORRAR
Private Sub btnBorrar_Click()
    Dim ultimaFila As Long
    Dim hojaDespachado As Worksheet
    Dim seleccion As Range
    Dim respuesta As VbMsgBoxResult

    Set hojaDespachado = ThisWorkbook.Sheets("PEDIDOS")

    ' Verificar si hay alguna fila seleccionada
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

' BORRAR FILAS VACIAS
Private Sub btnEliminarFilasVacias_Click()
    Dim hojaPalletizado As Worksheet
    Dim ultimaFila As Long
    Dim fila As Long
    Dim filasBorradas As Boolean

    Set hojaPalletizado = ThisWorkbook.Sheets("DESPACHADO")

    ultimaFila = hojaPalletizado.Cells(hojaPalletizado.Rows.Count, "A").End(xlUp).Row

    filasBorradas = False

    For fila = ultimaFila To 4 Step -1
        If hojaPalletizado.Cells(fila, 1).Value = "" Then
            ' La celda en la columna A está vacía, eliminar toda la fila
            hojaPalletizado.Rows(fila).Delete
            filasBorradas = True
        End If
    Next fila

    If Not filasBorradas Then
        MsgBox "No se encontraron filas vacías para borrar.", vbInformation
    End If
End Sub








