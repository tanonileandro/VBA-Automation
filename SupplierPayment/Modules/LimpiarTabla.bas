Attribute VB_Name = "LimpiarTabla"
Sub EliminarFilas()
    Dim wsDestino As Worksheet
    Dim tblDestino As ListObject
    Dim lastRow As Long

    On Error Resume Next
    Set wsDestino = ActiveSheet
    On Error GoTo 0

    If wsDestino Is Nothing Then
        MsgBox "No se pudo determinar la hoja activa.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects(1)
    On Error GoTo 0

    If tblDestino Is Nothing Then
        MsgBox "No se encontró una tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    lastRow = tblDestino.ListRows.Count

    If lastRow > 2 Then
        wsDestino.Rows("3:" & lastRow + 2).Delete
    Else
        MsgBox "No hay filas para eliminar en la hoja '" & wsDestino.Name & "'.", vbExclamation
    End If
End Sub

