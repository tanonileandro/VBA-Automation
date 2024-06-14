Attribute VB_Name = "ExportarAPagar"
Sub ExportarAPagar()
    On Error Resume Next
    
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tblOrigen As ListObject
    Dim tblDestino As ListObject
    Dim lastRowOrigen As Long
    Dim i As Long

    Set wsOrigen = ActiveSheet
    
    Set wsDestino = ThisWorkbook.Sheets("PAGADAS")
    
    On Error Resume Next
    Set tblOrigen = wsOrigen.ListObjects(1)
    On Error GoTo 0

    If tblOrigen Is Nothing Then
        MsgBox "No se encontró ninguna tabla en la hoja de origen.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("Tabla88")
    On Error GoTo 0

    If tblDestino Is Nothing Then
        MsgBox "La tabla de destino no se encontró en la hoja 'PAGADAS'.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    lastRowOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row

    For i = 1 To tblOrigen.ListRows.Count
        tblDestino.ListRows.Add
        ' Copiar los valores de la fila en la tabla de origen
        tblDestino.ListRows(tblDestino.ListRows.Count).Range.Offset(0, 1).Resize(1, tblOrigen.ListRows(i).Range.Columns.Count).Value = tblOrigen.ListRows(i).Range.Value
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Error al exportar a la hoja 'PAGADAS': " & Err.Description, vbExclamation
    Else
        MsgBox "Los datos se exportaron correctamente a la hoja 'PAGADAS'.", vbInformation
    End If

    On Error GoTo 0
End Sub

