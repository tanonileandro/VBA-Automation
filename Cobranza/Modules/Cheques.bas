Attribute VB_Name = "Cheques"
Sub ImportarChequesRechazados()
    Dim wsDestino As Worksheet
    Dim tablaDestino As ListObject
    Dim wsOrigen As Worksheet
    Dim tablaOrigen As ListObject
    Dim filaDestino As ListRow
    Dim i As Long, j As Long
    Dim hojaTabla As Variant
    Dim nombreHoja As String
    Dim nombreTabla As String
    Dim mesNombre As String

    ' Definir la hoja y la tabla de destino
    Set wsDestino = ThisWorkbook.Sheets("Historico Cheq Rechazados")
    Set tablaDestino = wsDestino.ListObjects("Cheques")
    
    ' Definir las hojas y tablas a importar
    Dim hojasTablas As Variant
    hojasTablas = Array( _
        Array("Carlos Cobo", "TablaCC"), _
        Array("Diego Picci", "TablaDP"), _
        Array("Horacio Schaad", "TablaHS"), _
        Array("Marcos Nadin", "TablaMN"), _
        Array("Pedro Iuorno", "TablaPI"), _
        Array("Rosario Pack", "TablaRP"), _
        Array("Embalajes", "TablaE") _
    )

    ' Recorrer cada hoja y tabla especificada
    For j = LBound(hojasTablas) To UBound(hojasTablas)
        nombreHoja = hojasTablas(j)(0)
        nombreTabla = hojasTablas(j)(1)
        
        ' Intentar establecer la hoja y la tabla de origen
        On Error Resume Next
        Set wsOrigen = ThisWorkbook.Sheets(nombreHoja)
        Set tablaOrigen = wsOrigen.ListObjects(nombreTabla)
        On Error GoTo 0
        
        If Not tablaOrigen Is Nothing Then
            ' Recorrer las filas de la tabla de origen
            For i = 1 To tablaOrigen.ListRows.count
                ' Verificar si el valor en la columna 3 es "AD"
                If tablaOrigen.ListColumns(3).DataBodyRange.Cells(i, 1).Value = "AD" Then
                    ' Agregar una nueva fila en la tabla de destino
                    Set filaDestino = tablaDestino.ListRows.Add
                    
                    ' Obtener el mes de la columna 9 y convertirlo a formato de tres letras
                    mesNombre = Format(tablaOrigen.ListColumns(9).DataBodyRange.Cells(i, 1).Value, "mmm")

                    ' Copiar datos a la nueva fila en la tabla de destino
                    With filaDestino
                        .Range(2).Value = mesNombre ' Nombre del mes con tres letras
                        .Range(3).Value = wsOrigen.Range("A2").Value ' Columna 1
                        .Range(4).Value = wsOrigen.Range("C1").Value ' Columna 2
                        .Range(5).Value = tablaOrigen.ListColumns(2).DataBodyRange.Cells(i, 1).Value ' Columna 3
                        .Range(6).Value = tablaOrigen.ListColumns(3).DataBodyRange.Cells(i, 1).Value ' Columna 4
                        .Range(7).Value = tablaOrigen.ListColumns(5).DataBodyRange.Cells(i, 1).Value ' Columna 5
                        .Range(8).Value = tablaOrigen.ListColumns(7).DataBodyRange.Cells(i, 1).Value ' Columna 6
                        .Range(9).Value = tablaOrigen.ListColumns(8).DataBodyRange.Cells(i, 1).Value ' Columna 7
                        .Range(10).Value = tablaOrigen.ListColumns(11).DataBodyRange.Cells(i, 1).Value ' Columna 8
                        .Range(13).Value = Now ' Fecha y hora actual
                    End With
                End If
            Next i
        End If
        
        ' Reiniciar las variables de origen
        Set wsOrigen = Nothing
        Set tablaOrigen = Nothing
    Next j

    ' Agregar una fila en blanco al final y colocar guiones en cada columna
    Set filaDestino = tablaDestino.ListRows.Add
    With filaDestino
        Dim col As Long
        For col = 1 To tablaDestino.ListColumns.count
            .Range(col).Value = "-"
        Next col
    End With

End Sub







