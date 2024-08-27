Attribute VB_Name = "ExportarHistorico"
Sub ExportarHistorico()
    On Error Resume Next
    
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tblOrigen As ListObject
    Dim tblDestino As ListObject
    Dim lastRowOrigen As Long
    Dim i As Long

    Set wsOrigen = ActiveSheet
    
    Set wsDestino = ThisWorkbook.Sheets("Historico Anual")
    
    On Error Resume Next
    Set tblOrigen = wsOrigen.ListObjects(1)
    On Error GoTo 0

    If tblOrigen Is Nothing Then
        MsgBox "No se encontró ninguna tabla en la hoja de origen.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("Tabla2")
    On Error GoTo 0

    If tblDestino Is Nothing Then
        MsgBox "La tabla de destino no se encontró en la hoja 'Historico Anual'.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    lastRowOrigen = wsOrigen.Cells(wsOrigen.Rows.count, "A").End(xlUp).Row

    For i = 1 To tblOrigen.ListRows.count
        tblDestino.ListRows.Add
        ' Copiar los valores de la fila en la tabla de origen
        tblDestino.ListRows(tblDestino.ListRows.count).Range.Offset(0, 1).Resize(1, tblOrigen.ListRows(i).Range.Columns.count).Value = tblOrigen.ListRows(i).Range.Value
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Error al exportar a la hoja 'Historico Anual': " & Err.Description, vbExclamation
    Else
        MsgBox "Los datos se exportaron correctamente a la hoja 'Historico Anual'.", vbInformation
    End If

    On Error GoTo 0
End Sub
