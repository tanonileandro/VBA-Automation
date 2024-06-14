Attribute VB_Name = "LimpiarTablaProveedores"
Sub LimpiarContenidoProveedores()
    Dim wsProveedores As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set wsProveedores = ThisWorkbook.Sheets("Proveedores")
    On Error GoTo 0

    If wsProveedores Is Nothing Then
        MsgBox "No se encontró la hoja 'Proveedores'.", vbExclamation
        Exit Sub
    End If

    lastRow = wsProveedores.Cells(wsProveedores.Rows.Count, "A").End(xlUp).Row

    If lastRow > 3 Then
        wsProveedores.Range("A3:A" & lastRow - 1).ClearContents
        wsProveedores.Range("C3:C" & lastRow - 1).ClearContents
        wsProveedores.Range("D3:D" & lastRow - 1).ClearContents
        wsProveedores.Range("E3:E" & lastRow - 1).ClearContents
        wsProveedores.Range("G3:G" & lastRow - 1).ClearContents
        wsProveedores.Range("H3:H" & lastRow - 1).ClearContents

        MsgBox "El contenido de las filas se limpió correctamente en la hoja 'Proveedores'.", vbInformation
    Else
        MsgBox "No hay filas con contenido para limpiar en la hoja 'Proveedores'.", vbExclamation
    End If
End Sub


