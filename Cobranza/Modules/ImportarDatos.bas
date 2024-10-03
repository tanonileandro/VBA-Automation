Attribute VB_Name = "ImportarDatos"
Sub ImportarDatos()

    Application.ScreenUpdating = False
  
    Dim respuesta As VbMsgBoxResult
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tablaOrigen As ListObject
    Dim tablaDestino As ListObject
    Dim UltimaFila As Long
    Dim i As Long
    Dim data As Variant
    
    ' Definir las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("E-Zential")
    Set wsDestino = ThisWorkbook.Sheets("COBRANZA TOTAL")
    
    ' Definir la tabla de origen (TablaCC)
    Set tablaOrigen = wsOrigen.ListObjects("Tabla1")
    
    ' Definir la tabla de destino (Cobranza)
    Set tablaDestino = wsDestino.ListObjects("Tabla2")
    
    ' Preguntar al usuario si desea proceder
    respuesta = MsgBox("¿Está seguro de Importar? Se eliminarán los datos actuales de esta planilla.", vbYesNo + vbQuestion, "Confirmar Exportación")
    
    If respuesta = vbNo Then
        Exit Sub ' Salir de la macro si el usuario elige No
    End If
    
    ' Quitar filtros en la tabla de destino, si están activos
    QuitarFiltros wsDestino, "Tabla2"
    Vaciar

    ' Leer todos los datos de la tabla de origen en un array
    data = tablaOrigen.DataBodyRange.Value
    
    ' Agregar las nuevas filas a la tabla de destino
    For i = 1 To UBound(data, 1)
        With tablaDestino.ListRows.Add
            .Range(1, 1).Value = data(i, 2)
            .Range(1, 2).Value = data(i, 3)
            .Range(1, 3).Value = data(i, 5)
            .Range(1, 4).Value = data(i, 16)
            .Range(1, 5).Value = data(i, 17)
            .Range(1, 6).Value = data(i, 18)
            .Range(1, 7).Value = data(i, 23)
            .Range(1, 8).Value = data(i, 24)
            .Range(1, 9).Value = data(i, 19)
            .Range(1, 11).Value = data(i, 21)
            .Range(1, 12).Value = data(i, 32)
        End With
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub QuitarFiltros(ws As Worksheet, tableName As String)
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If Not tbl Is Nothing Then
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    End If
End Sub

