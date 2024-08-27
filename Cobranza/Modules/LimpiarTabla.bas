Attribute VB_Name = "LimpiarTabla"
Sub VaciarTabla()

    Dim Hoja As Worksheet
    Dim tabla As ListObject
    
    ' Establecer la hoja de trabajo activa
    Set Hoja = ActiveSheet
    
    ' Asumir que hay solo una tabla en la hoja activa
    On Error Resume Next
    Set tabla = Hoja.ListObjects(1)
    On Error GoTo 0
    
    If tabla Is Nothing Then
        MsgBox "No se encontró ninguna tabla en la hoja activa.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si la tabla está vacía
    If tabla.ListRows.Count = 0 Then
        MsgBox "La tabla ya está vacía.", vbInformation
        Exit Sub
    End If
    
    ' Eliminar todas las filas de la tabla
    tabla.DataBodyRange.Rows.Delete
    
End Sub

