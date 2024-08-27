Attribute VB_Name = "InsertarSem"
Private Sub BotonActualizarSemanas_Click()
    ' Llamar al procedimiento para actualizar las semanas
    ActualizarSemanas
End Sub

Sub ActualizarSemanas()
    Dim rng As Range
    Dim cell As Range
    Dim fecha As Date
    Dim semana As Variant
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colFecha As ListColumn
    Dim colSemana As ListColumn
    
    ' Referenciar la hoja activa
    Set ws = ActiveSheet
    
    ' Intentar obtener la primera tabla en la hoja activa
    On Error Resume Next
    Set tbl = ws.ListObjects(1)
    On Error GoTo 0
    
    ' Salir si no hay tabla encontrada
    If tbl Is Nothing Then Exit Sub
    
    ' Buscar la columna de fecha y la columna de semana por sus nombres
    Set colFecha = tbl.ListColumns("Fecha Vto")
    Set colSemana = tbl.ListColumns("Sem")
    
    ' Salir si no se encontraron las columnas
    If colFecha Is Nothing Or colSemana Is Nothing Then Exit Sub
    
    ' Definir el rango de la columna de fecha en la tabla
    Set rng = colFecha.DataBodyRange
    
    Application.EnableEvents = False  ' Desactivar eventos para evitar bucles infinitos
    
    For Each cell In rng
        If IsDate(cell.Value) Then
            ' Convertir el valor de la celda a fecha
            fecha = CDate(cell.Value)
            
            ' Calcular el número de semana usando ISO.SEMANA.NUMERO
            semana = Application.WorksheetFunction.IsoWeekNum(fecha)
            
            ' Escribir el número de semana en la columna de semana
            cell.Offset(0, colSemana.Index - colFecha.Index).Value = semana
        End If
    Next cell
    
    Application.EnableEvents = True  ' Habilitar eventos nuevamente
End Sub






    

