Attribute VB_Name = "EstadoSem"
Sub BorrarImportar()

    Vaciar
    ImportarHistoricoClienteDesdeRotulo
    ActualizarVendedor

End Sub
Sub ImportarHistoricoClienteDesdeRotulo()

    Application.ScreenUpdating = False
    
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tablaOrigen As ListObject
    Dim tablaDestino As ListObject
    Dim nuevaFila As ListRow
    Dim nombresDict As Object
    Dim sumaColumna14Dict As Object
    Dim nombre As Variant
    Dim i As Long
    Dim sheetName As String
    Dim letraFiltro As String
    Dim tableName As String
    Dim filaDestino As Long
    Dim valorColumnaA As String
    Dim valorLimite As Double
    
    ' Obtener el nombre de la hoja y la letra desde la celda I2 y J2 de la hoja "Estado Sem."
    sheetName = ThisWorkbook.Sheets("Estado Sem.").Range("I2").Value
    letraFiltro = ThisWorkbook.Sheets("Estado Sem.").Range("J2").Value
    
    ' Determinar el nombre de la tabla en función de la hoja seleccionada
    Select Case sheetName
        Case "Carlos Cobo": tableName = "TablaCC"
        Case "Diego Picci": tableName = "TablaDP"
        Case "Horacio Schaad": tableName = "TablaHS"
        Case "Marcos Nadin": tableName = "TablaMN"
        Case "Pedro Iuorno": tableName = "TablaPI"
        Case "Rosario Pack": tableName = "TablaRP"
        Case "Embalajes": tableName = "TablaE"
        Case Else
            MsgBox "El nombre de la hoja no coincide con ninguna de las tablas definidas: " & sheetName, vbExclamation
            Exit Sub
    End Select
    
    ' Crear diccionarios para evitar duplicados y sumar columnas 7 y 14
    Set nombresDict = CreateObject("Scripting.Dictionary")
    Set sumaColumna14Dict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wsOrigen = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If wsOrigen Is Nothing Then
        MsgBox "La hoja '" & sheetName & "' no se encontró en el libro.", vbExclamation
        Exit Sub
    End If
    
    ' Definir la tabla de origen en la hoja de origen
    On Error Resume Next
    Set tablaOrigen = wsOrigen.ListObjects(tableName)
    On Error GoTo 0
    
    If tablaOrigen Is Nothing Then
        MsgBox "No se encontró la tabla '" & tableName & "' en la hoja '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener el valor de la celda C1 para el límite
    valorLimite = wsOrigen.Range("C1").Value
    
    ' Definir la tabla de destino "Rotulo" en la hoja "Estado Sem."
    Set wsDestino = ThisWorkbook.Sheets("Estado Sem.")
    Set tablaDestino = wsDestino.ListObjects("Rotulo")
    
    ' Recorrer la columna 2 (nombre) y sumar los valores de las columnas 7 y 14 solo si la letra coincide
    For i = 1 To tablaOrigen.ListRows.count
        ' Solo procesar si el valor en la columna 7 es mayor a 0 y la columna 12 no es mayor que valorLimite
        If tablaOrigen.ListColumns(7).DataBodyRange.Cells(i, 1).Value > 0 And _
           tablaOrigen.ListColumns(12).DataBodyRange.Cells(i, 1).Value <= valorLimite Then
            
            nombre = tablaOrigen.ListColumns(2).DataBodyRange.Cells(i, 1).Value
            valorColumnaA = tablaOrigen.ListColumns(1).DataBodyRange.Cells(i, 1).Value
            Dim valorLetra As String
            valorLetra = tablaOrigen.ListColumns(4).DataBodyRange.Cells(i, 1).Value ' Letra de la columna 4
            
            ' Verificar que el valor de la letra coincide con la celda J2
            If valorLetra = letraFiltro Then
                ' Si el nombre no está en el diccionario, agregarlo y sumar columnas 7 y 14
                If Not nombresDict.Exists(nombre) Then
                    nombresDict.Add nombre, tablaOrigen.ListColumns(7).DataBodyRange.Cells(i, 1).Value
                    sumaColumna14Dict.Add nombre, tablaOrigen.ListColumns(14).DataBodyRange.Cells(i, 1).Value
                Else
                    ' Si el nombre ya existe, sumar los valores de las columnas 7 y 14
                    nombresDict(nombre) = nombresDict(nombre) + tablaOrigen.ListColumns(7).DataBodyRange.Cells(i, 1).Value
                    sumaColumna14Dict(nombre) = sumaColumna14Dict(nombre) + tablaOrigen.ListColumns(14).DataBodyRange.Cells(i, 1).Value
                End If
            End If
        End If
    Next i
    
    ' Agregar los resultados únicos a la tabla de destino (sin error al agregar filas)
    For Each nombre In nombresDict.Keys
        ' Obtener el valor correspondiente de la columna A para el nombre actual
        For i = 1 To tablaOrigen.ListRows.count
            If tablaOrigen.ListColumns(2).DataBodyRange.Cells(i, 1).Value = nombre Then
                valorColumnaA = tablaOrigen.ListColumns(1).DataBodyRange.Cells(i, 1).Value
                Exit For ' Salir del bucle una vez que encontramos la coincidencia
            End If
        Next i
        
        Set nuevaFila = tablaDestino.ListRows.Add ' Agrega fila automáticamente
        filaDestino = nuevaFila.Index
        
        ' Copiar datos a la nueva fila en la tabla de destino
        With tablaDestino.ListRows(filaDestino)
            .Range(1).Value = wsOrigen.Range("A2").Value ' Otro valor de origen
            .Range(2).Value = letraFiltro ' Colocar la letra A o B en la columna 2
            .Range(4).Value = nombre ' Columna 2 (nombre)
            .Range(3).Value = valorColumnaA ' Valor correspondiente de la columna A
            .Range(5).Value = nombresDict(nombre) ' Sumatoria de columna 7 por nombre
            .Range(6).Value = sumaColumna14Dict(nombre) ' Sumatoria de columna 14 por nombre
        End With
    Next nombre
    
    Application.ScreenUpdating = True

End Sub


Sub SeleccionarRango()
    Dim wsRotulo As Worksheet
    Dim UltimaFila As Long
    
    ' Hoja de cálculo a tomar
    Set wsRotulo = Worksheets("Estado Sem.")
    
    ' Encuentra la última fila ocupada en la columna A
    UltimaFila = wsRotulo.Cells(Rows.count, "A").End(xlUp).Row
    
    ' Seleccionar el rango desde A1 hasta la última fila del rótulo
    Dim RangoInforme As Range
    Set RangoInforme = wsRotulo.Range("A1:G" & UltimaFila)
    RangoInforme.Select
    
    ' Agregar borde alrededor del rango seleccionado
    RangoInforme.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(0, 0, 0) ' Estilo de línea continuo, grosor mediano, color negro
End Sub



