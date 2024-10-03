Attribute VB_Name = "LimpiarTabla"
Sub VaciarTabla()

    Dim Hoja As Worksheet
    Dim tabla As ListObject
    Dim respuesta As VbMsgBoxResult
    
    respuesta = MsgBox("¿Está seguro de que quiere vaciar la planilla? Se perderán los datos anteriores.", vbYesNo + vbQuestion, "Confirmar Exportación")
    
    If respuesta = vbNo Then
        Exit Sub ' Salir de la macro si el usuario elige No
    End If
    
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
    If tabla.ListRows.count = 0 Then
        MsgBox "La tabla ya está vacía.", vbInformation
        Exit Sub
    End If
    
    ' Eliminar todas las filas de la tabla
    tabla.DataBodyRange.Rows.Delete
    
End Sub

Sub Vaciar()

    Dim Hoja As Worksheet
    Dim tabla As ListObject
    Dim respuesta As VbMsgBoxResult
    
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
    If tabla.ListRows.count = 0 Then
        MsgBox "La tabla ya está vacía.", vbInformation
        Exit Sub
    End If
    
    ' Eliminar todas las filas de la tabla
    tabla.DataBodyRange.Rows.Delete
    
End Sub
