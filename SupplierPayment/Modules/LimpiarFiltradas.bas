Attribute VB_Name = "LimpiarFiltradas"
Sub LimpiarFiltradoVisible()
    Dim rngVisible As Range
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    
    ' Definir la hoja de trabajo activa
    Set ws = ActiveSheet
    
    ' Comprobar si hay una tabla en la hoja activa
    If ws.ListObjects.Count = 0 Then
        MsgBox "No hay ninguna tabla en esta hoja.", vbExclamation
        Exit Sub
    End If
    
    ' Utilizar la primera tabla encontrada en la hoja activa
    Set tbl = ws.ListObjects(1)
    
    ' Filtrar y seleccionar solo las filas visibles
    On Error Resume Next
    Set rngVisible = tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not rngVisible Is Nothing Then
        ' Desagrupar las filas antes de borrar
        ws.Outline.ShowLevels RowLevels:=1
        ' Eliminar las filas visibles
        rngVisible.Rows.Delete
    Else
        MsgBox "No hay filas visibles para limpiar.", vbExclamation
    End If
End Sub
