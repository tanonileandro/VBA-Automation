Attribute VB_Name = "LimpiarTabla"
Sub VaciarTabla()

    Dim Hoja As Worksheet
    Dim tabla As ListObject
    Dim respuesta As VbMsgBoxResult
    
    respuesta = MsgBox("�Est� seguro de que quiere vaciar la planilla? Se perder�n los datos anteriores.", vbYesNo + vbQuestion, "Confirmar Exportaci�n")
    
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
        MsgBox "No se encontr� ninguna tabla en la hoja activa.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si la tabla est� vac�a
    If tabla.ListRows.count = 0 Then
        MsgBox "La tabla ya est� vac�a.", vbInformation
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
        MsgBox "No se encontr� ninguna tabla en la hoja activa.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si la tabla est� vac�a
    If tabla.ListRows.count = 0 Then
        MsgBox "La tabla ya est� vac�a.", vbInformation
        Exit Sub
    End If
    
    ' Eliminar todas las filas de la tabla
    tabla.DataBodyRange.Rows.Delete
    
End Sub
