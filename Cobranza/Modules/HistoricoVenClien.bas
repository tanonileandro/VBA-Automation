Attribute VB_Name = "HistoricoVenClien"
Sub ActualizarImportarVendedor()
    ' Actualizar las sumatorias primero
    ActSemanaSumatorias "Carlos Cobo", "TablaCC"
    ActSemanaSumatorias "Diego Picci", "TablaDP"
    ActSemanaSumatorias "Horacio Schaad", "TablaHS"
    ActSemanaSumatorias "Marcos Nadin", "TablaMN"
    ActSemanaSumatorias "Pedro Iuorno", "TablaPI"
    ActSemanaSumatorias "Rosario Pack", "TablaRP"
    ActSemanaSumatorias "Embalajes", "TablaE"
    
    ' Importar todos los datos de los vendedores
    ImportarHistoricoVendedor "Carlos Cobo", "TablaCC"
    ImportarHistoricoVendedor "Diego Picci", "TablaDP"
    ImportarHistoricoVendedor "Horacio Schaad", "TablaHS"
    ImportarHistoricoVendedor "Marcos Nadin", "TablaMN"
    ImportarHistoricoVendedor "Pedro Iuorno", "TablaPI"
    ImportarHistoricoVendedor "Rosario Pack", "TablaRP"
    ImportarHistoricoVendedor "Embalajes", "TablaE"
    
    ' Mostrar el mensaje de confirmación al final de la importación de todas las tablas
    MsgBox "Todos los datos han sido copiados correctamente al Historico Vendedor.", vbInformation
End Sub

Sub ActualizarImportarCliente()
    ' Actualizar las sumatorias primero
    ActSemanaSumatorias "Carlos Cobo", "TablaCC"
    ActSemanaSumatorias "Diego Picci", "TablaDP"
    ActSemanaSumatorias "Horacio Schaad", "TablaHS"
    ActSemanaSumatorias "Marcos Nadin", "TablaMN"
    ActSemanaSumatorias "Pedro Iuorno", "TablaPI"
    ActSemanaSumatorias "Rosario Pack", "TablaRP"
    ActSemanaSumatorias "Embalajes", "TablaE"
    
    ' Importar todos los datos de los vendedores
    ImportarHistoricoCliente "Carlos Cobo", "TablaCC"
    ImportarHistoricoCliente "Diego Picci", "TablaDP"
    ImportarHistoricoCliente "Horacio Schaad", "TablaHS"
    ImportarHistoricoCliente "Marcos Nadin", "TablaMN"
    ImportarHistoricoCliente "Pedro Iuorno", "TablaPI"
    ImportarHistoricoCliente "Rosario Pack", "TablaRP"
    ImportarHistoricoCliente "Embalajes", "TablaE"
    
    ' Mostrar el mensaje de confirmación al final de la importación de todas las tablas
    MsgBox "Todos los datos han sido copiados correctamente al Historico Vendedor.", vbInformation
End Sub
Sub ActSemanaSumatorias(nombreHoja As String, nombreTabla As String)

    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim numeroSemana As Integer
    Dim Hoja As Worksheet
    Dim tabla As ListObject
    Dim i As Long
    Dim sumaColumnaG As Double
    Dim sumaColumnaG_A As Double
    Dim sumaColumnaG_B As Double
    Dim sumaColumnaN As Double
    Dim sumaColumnaN_A As Double
    Dim sumaColumnaN_B As Double
    
    ' Establecer la hoja de trabajo
    On Error Resume Next
    Set Hoja = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo 0
    
    If Hoja Is Nothing Then
        MsgBox "No se encontró la hoja: " & nombreHoja, vbExclamation
        Exit Sub
    End If
    
    ' Establecer la tabla
    On Error Resume Next
    Set tabla = Hoja.ListObjects(nombreTabla)
    On Error GoTo 0
    
    If tabla Is Nothing Then
        MsgBox "No se encontró la tabla: " & nombreTabla, vbExclamation
        Exit Sub
    End If

    ' Calcular la fecha de inicio y fin de la semana
    fechaInicio = Date - (Weekday(Date, vbMonday) - 1)
    fechaFin = fechaInicio + 6
    numeroSemana = Application.WorksheetFunction.WeekNum(Date, vbMonday)
    
    ' Escribir las fechas y el número de semana en las celdas A1 y A2
    With Hoja
        .Range("A1").Value = "Semana " & numeroSemana
        .Range("A2").Value = Format(fechaInicio, "dd-mm") & " al " & Format(fechaFin, "dd-mm")
        .Range("C1").Value = numeroSemana
    End With

    ' Inicializar sumatorias
    sumaColumnaG = 0
    sumaColumnaG_A = 0
    sumaColumnaG_B = 0
    sumaColumnaN = 0
    sumaColumnaN_A = 0
    sumaColumnaN_B = 0
    
    ' Calcular sumatorias
    For i = 1 To tabla.ListRows.count
      
        If tabla.DataBodyRange(i, 7).Value >= 0 And _
           tabla.DataBodyRange(i, 12).Value <= numeroSemana Then
            sumaColumnaG = sumaColumnaG + tabla.DataBodyRange(i, 7).Value
        End If
        
        ' Sumar según el valor de la columna 4 (Columna D)
        If tabla.DataBodyRange(i, 4).Value = "A" Then
            If tabla.DataBodyRange(i, 7).Value >= 0 And _
               tabla.DataBodyRange(i, 12).Value <= numeroSemana Then
                sumaColumnaG_A = sumaColumnaG_A + tabla.DataBodyRange(i, 7).Value
            End If
            sumaColumnaN_A = sumaColumnaN_A + tabla.DataBodyRange(i, 14).Value
        ElseIf tabla.DataBodyRange(i, 4).Value = "B" Then
            If tabla.DataBodyRange(i, 7).Value >= 0 And _
               tabla.DataBodyRange(i, 12).Value <= numeroSemana Then
                sumaColumnaG_B = sumaColumnaG_B + tabla.DataBodyRange(i, 7).Value
            End If
            sumaColumnaN_B = sumaColumnaN_B + tabla.DataBodyRange(i, 14).Value
        End If

        ' Sumar la columna 14 (Columna N)
        sumaColumnaN = sumaColumnaN + tabla.DataBodyRange(i, 14).Value
    Next i
    
    ' Pegar las sumatorias en las celdas correspondientes
    Hoja.Range("M2").Value = sumaColumnaG
    Hoja.Range("F1").Value = sumaColumnaG_A
    Hoja.Range("F2").Value = sumaColumnaG_B
    Hoja.Range("M3").Value = sumaColumnaN
    Hoja.Range("H1").Value = sumaColumnaN_A
    Hoja.Range("H2").Value = sumaColumnaN_B
    
    ' Calcular y mostrar la diferencia en M4
    Hoja.Range("M4").Value = sumaColumnaG - sumaColumnaN
     
End Sub

Sub ImportarHistoricoVendedor(sheetName As String, tableName As String)
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tablaOrigen As ListObject
    Dim tablaDestino As ListObject
    Dim nuevaFila As ListRow
    
    On Error Resume Next
    Set wsOrigen = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If wsOrigen Is Nothing Then
        MsgBox "La hoja '" & sheetName & "' no se encontró en el libro.", vbExclamation
        Exit Sub
    End If
    
    ' Definir la hoja destino "Historico Vendedor"
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Historico Vendedor")
    On Error GoTo 0
    
    If wsDestino Is Nothing Then
        ' Si la hoja destino no existe, crear una nueva hoja
        Set wsDestino = ThisWorkbook.Sheets.Add
        wsDestino.Name = "Historico Vendedor"
    End If
    
    ' Definir la tabla de origen en la hoja de origen
    On Error Resume Next
    Set tablaOrigen = wsOrigen.ListObjects(tableName)
    On Error GoTo 0
    
    If tablaOrigen Is Nothing Then
        MsgBox "No se encontró la tabla '" & tableName & "' en la hoja '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Definir la tabla de destino "Historico" en la hoja destino
    On Error Resume Next
    Set tablaDestino = wsDestino.ListObjects("Historico")
    On Error GoTo 0
    
    If tablaDestino Is Nothing Then
        Set tablaDestino = wsDestino.ListObjects.Add(xlSrcRange, wsDestino.Range("A1"), , xlYes)
        tablaDestino.Name = "Historico"
    End If
    
    ' Agregar una nueva fila en la tabla de destino
    Set nuevaFila = tablaDestino.ListRows.Add
    
    ' Copiar datos desde la tabla de origen a la nueva fila en la tabla de destino
    With nuevaFila
        .Range(1).Value = wsOrigen.Range("A2").Value
        .Range(2).Value = wsOrigen.Range("C1").Value
        .Range(3).Value = sheetName ' Nombre de la hoja de origen
        .Range(4).Value = wsOrigen.Range("M2").Value
        .Range(5).Value = wsOrigen.Range("F1").Value
        .Range(6).Value = wsOrigen.Range("F2").Value
        .Range(7).Value = wsOrigen.Range("M3").Value
        .Range(9).Value = Now ' Fecha y hora actual
    End With
    
End Sub

Sub ImportarHistoricoCliente(sheetName As String, tableName As String)
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tablaOrigen As ListObject
    Dim tablaDestino As ListObject
    Dim nuevaFila As ListRow
    Dim nombresDict As Object
    Dim sumaColumna14Dict As Object ' Diccionario para la sumatoria de la columna 14
    Dim nombre As Variant
    Dim i As Long
    Dim sumaColumna7 As Double
    Dim sumaColumna14 As Double
    Dim filaDestino As Long
    
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
    
    ' Definir la hoja destino "Historico Cliente"
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Historico Cliente")
    On Error GoTo 0
    
    If wsDestino Is Nothing Then
        ' Si la hoja destino no existe, crear una nueva hoja
        Set wsDestino = ThisWorkbook.Sheets.Add
        wsDestino.Name = "Historico Cliente"
    End If
    
    ' Definir la tabla de origen en la hoja de origen
    On Error Resume Next
    Set tablaOrigen = wsOrigen.ListObjects(tableName)
    On Error GoTo 0
    
    If tablaOrigen Is Nothing Then
        MsgBox "No se encontró la tabla '" & tableName & "' en la hoja '" & sheetName & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Definir la tabla de destino "HistoricoClientes" en la hoja destino
    On Error Resume Next
    Set tablaDestino = wsDestino.ListObjects("HistoricoClientes")
    On Error GoTo 0
    
    If tablaDestino Is Nothing Then
        Set tablaDestino = wsDestino.ListObjects.Add(xlSrcRange, wsDestino.Range("A1"), , xlYes)
        tablaDestino.Name = "HistoricoClientes"
    End If
    
    ' Recorrer la columna 2 (nombre) y sumar los valores de las columnas 7 y 14
    For i = 1 To tablaOrigen.ListRows.count
        nombre = tablaOrigen.ListColumns(2).DataBodyRange.Cells(i, 1).Value
        
        ' Si el nombre no está en el diccionario, agregarlo y sumar columnas 7 y 14
        If Not nombresDict.Exists(nombre) Then
            nombresDict.Add nombre, tablaOrigen.ListColumns(7).DataBodyRange.Cells(i, 1).Value
            sumaColumna14Dict.Add nombre, tablaOrigen.ListColumns(14).DataBodyRange.Cells(i, 1).Value
        Else
            ' Si el nombre ya existe, sumar los valores de las columnas 7 y 14
            nombresDict(nombre) = nombresDict(nombre) + tablaOrigen.ListColumns(7).DataBodyRange.Cells(i, 1).Value
            sumaColumna14Dict(nombre) = sumaColumna14Dict(nombre) + tablaOrigen.ListColumns(14).DataBodyRange.Cells(i, 1).Value
        End If
    Next i
    
    ' Agregar los resultados únicos a la tabla de destino
    For Each nombre In nombresDict.Keys ' nombre ahora es Variant
        Set nuevaFila = tablaDestino.ListRows.Add
        filaDestino = nuevaFila.Index
        
        ' Copiar datos a la nueva fila en la tabla de destino
        With tablaDestino.ListRows(filaDestino)
            .Range(1).Value = wsOrigen.Range("A2").Value
            .Range(2).Value = wsOrigen.Range("C1").Value
            .Range(3).Value = sheetName ' Nombre de la hoja de origen
            .Range(4).Value = nombre ' Columna 2 (nombre)
            .Range(5).Value = nombresDict(nombre) ' Sumatoria de columna 7 por nombre
            .Range(6).Value = sumaColumna14Dict(nombre) ' Sumatoria de columna 14 por nombre
            .Range(8).Value = Now ' Fecha y hora actual
        End With
    Next nombre
    
End Sub



