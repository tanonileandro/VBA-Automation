Attribute VB_Name = "ExportarVendedores"
Sub ExportarTodos()
    Dim respuesta As VbMsgBoxResult
    
    ' Preguntar al usuario si desea proceder
    respuesta = MsgBox("¿Está seguro de exportar? Se eliminarán los datos de la semana anterior.", vbYesNo + vbQuestion, "Confirmar Exportación")
    
    If respuesta = vbNo Then
        Exit Sub ' Salir de la macro si el usuario elige No
    End If
    
    ' Limpiar todas las tablas antes de exportar
    LimpiarTabla "Carlos Cobo", "TablaCC"
    LimpiarTabla "Diego Picci", "TablaDP"
    LimpiarTabla "Horacio Schaad", "TablaHS"
    LimpiarTabla "Marcos Nadin", "TablaMN"
    LimpiarTabla "Pedro Iuorno", "TablaPI"
    LimpiarTabla "Rosario Pack", "TablaRP"
    LimpiarTabla "Embalajes", "TablaE"
    
    ' Proceder con la exportación
    ExportarMasivo "Carlos Cobo", "TablaCC"
    ExportarMasivo "Diego Picci", "TablaDP"
    ExportarMasivo "Horacio Schaad", "TablaHS"
    ExportarMasivo "Marcos Nadin", "TablaMN"
    ExportarMasivo "Pedro Iuorno", "TablaPI"
    ExportarMasivo "Rosario Pack", "TablaRP"
    ExportarMasivo "Embalajes", "TablaE"
    
    ' Depurar y actualizar la semana en todas las tablas
    DepurarYActualizar "Carlos Cobo", "TablaCC"
    DepurarYActualizar "Diego Picci", "TablaDP"
    DepurarYActualizar "Horacio Schaad", "TablaHS"
    DepurarYActualizar "Marcos Nadin", "TablaMN"
    DepurarYActualizar "Pedro Iuorno", "TablaPI"
    DepurarYActualizar "Rosario Pack", "TablaRP"
    DepurarYActualizar "Embalajes", "TablaE"
    
    MsgBox "Proceso completado satisfactoriamente.", vbInformation
    
End Sub

Private Sub LimpiarTabla(sheetName As String, tableName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            ' Eliminar todas las filas de la tabla
            If tbl.ListRows.count > 0 Then
                tbl.DataBodyRange.Delete
            End If
        Else
            MsgBox "No se encontró la tabla " & tableName & " en la hoja " & sheetName, vbExclamation
        End If
    Else
        MsgBox "No se encontró la hoja " & sheetName, vbExclamation
    End If
End Sub

Sub ExportarMasivo(sheetName As String, tableName As String)
    Application.ScreenUpdating = False

    Dim wsProveedores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaProveedores As Long
    Dim tblDestino As ListObject
    Dim i As Long
    Dim columnasAColar As Variant
    Dim j As Long
    Dim valor As Variant

    Set wsProveedores = ThisWorkbook.Sheets("COBRANZA TOTAL")

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ultimaFilaProveedores = wsProveedores.Cells(wsProveedores.Rows.count, "C").End(xlUp).Row

    ' Verificar la presencia de valores N/A en la columna de interés
    hayNAs = False
    For i = 5 To ultimaFilaProveedores
        If IsError(wsProveedores.Cells(i, 5).Value) Then
            If wsProveedores.Cells(i, 5).Value = CVErr(xlErrNA) Then
                hayNAs = True
                Exit For
            End If
        End If
    Next i

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects(tableName)
    On Error GoTo 0
    
    If tblDestino Is Nothing Then
        MsgBox "No se encontró la tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    ' Quitar filtros en la tabla de destino
    QuitarFiltros wsDestino, tableName

    ' Columnas que queremos copiar (en el orden deseado)
    columnasAColar = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15)

    For i = 5 To ultimaFilaProveedores
        If Not IsEmpty(wsProveedores.Cells(i, 5).Value) Then
            ' Compara los valores de la columna C después de aplicar Trim y UCase
            If UCase(Trim(wsProveedores.Cells(i, 3).Value)) = UCase(sheetName) Then ' Comparación con el nombre de la hoja
                tblDestino.ListRows.Add
                ' Copiar solo las columnas específicas
                For j = LBound(columnasAColar) To UBound(columnasAColar)
                    valor = wsProveedores.Cells(i, columnasAColar(j)).Value
                    tblDestino.ListRows(tblDestino.ListRows.count).Range.Cells(1, j + 1).Value = valor
                Next j
            End If
        End If
    Next i

    ' Llamar a la función para poner "B" en la columna 4 si hay "PS" en la columna 3
    ActualizarColumnaB tblDestino

    Application.ScreenUpdating = True
End Sub

Private Sub ActualizarColumnaB(tbl As ListObject)
    Dim i As Long
    Dim fila As ListRow
    For Each fila In tbl.ListRows
        If UCase(Trim(fila.Range.Cells(1, 3).Value)) = "PS" Then
            fila.Range.Cells(1, 4).Value = "B"
        End If
    Next fila
End Sub

Private Sub QuitarFiltros(ws As Worksheet, tableName As String)
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If Not tbl Is Nothing Then
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    End If
End Sub

Private Sub DepurarYActualizar(sheetName As String, tableName As String)

    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim numeroSemana As Integer
    Dim Hoja As Worksheet
    Dim i As Long
    Dim tabla As ListObject
    Dim columnaLIndex As Integer
    Dim columna14Index As Integer ' Índice para la columna 14
    
    ' Establecer la hoja de trabajo activa
    Set Hoja = ThisWorkbook.Sheets(sheetName)
    
    ' Calcular la fecha de inicio de la semana (lunes)
    fechaInicio = Date - (Weekday(Date, vbMonday) - 1)
    
    ' Calcular la fecha de fin de la semana (domingo)
    fechaFin = fechaInicio + 6
    
    ' Calcular el número de semana
    numeroSemana = Application.WorksheetFunction.WeekNum(Date, vbMonday)
    
    ' Escribir las fechas y el número de semana en las celdas A1 y A2 de la hoja activa
    With Hoja
        .Range("A1").Value = "Semana " & numeroSemana
        .Range("A2").Value = Format(fechaInicio, "dd-mm") & " al " & Format(fechaFin, "dd-mm")
        .Range("B3").Value = Format(fechaInicio, "dd-mm") & " al " & Format(fechaFin, "dd-mm")
        .Range("C1").Value = numeroSemana
    End With
    
    ' Asumir que hay solo una tabla en la hoja activa con el nombre tableName
    On Error Resume Next
    Set tabla = Hoja.ListObjects(tableName)
    On Error GoTo 0
    
    If tabla Is Nothing Then
        MsgBox "No se encontró la tabla '" & tableName & "' en la hoja '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Suponiendo que la columna L corresponde a la columna 12 dentro de la tabla
    columnaLIndex = 12
    columna14Index = 14
    
    ' Verificar si la columna L está dentro del rango de columnas de la tabla
    If columnaLIndex > tabla.ListColumns.count Then
        MsgBox "La tabla '" & tableName & "' en la hoja '" & sheetName & "' no tiene suficientes columnas para incluir la columna L.", vbExclamation
        Exit Sub
    End If
    
    ' Eliminar filas donde el valor en la columna L sea mayor que numeroSemana
    'For i = tabla.ListRows.Count To 1 Step -1
    '    If IsNumeric(tabla.DataBodyRange(i, columnaLIndex).Value) Then
    '        If tabla.DataBodyRange(i, columnaLIndex).Value > numeroSemana Then
    '            tabla.ListRows(i).Delete
    '        End If
    '    End If
    'Next i
    
    ' Establecer el valor 0 en cada celda de la columna 14
    For i = 1 To tabla.ListRows.count
        tabla.DataBodyRange(i, columna14Index).Value = 0
    Next i

End Sub



