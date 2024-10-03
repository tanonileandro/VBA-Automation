Attribute VB_Name = "Mensual"
Sub ImportarDatosMensual()

    Dim wsMensual As Worksheet
    Dim wsCarlosCobo As Worksheet
    Dim wsEstadoSem As Worksheet
    Dim tblMensual As ListObject
    Dim nuevaFila As ListRow

    ' Definir las hojas
    Set wsMensual = ThisWorkbook.Sheets("% Mensual")
    Set wsCarlosCobo = ThisWorkbook.Sheets("Carlos Cobo")
    Set wsEstadoSem = ThisWorkbook.Sheets("Estado Sem.")

    ' Definir la tabla en la hoja % Mensual
    Set tblMensual = wsMensual.ListObjects("Mensual") ' Asegúrate de que este nombre coincide

    ' Agregar una nueva fila a la tabla
    Set nuevaFila = tblMensual.ListRows.Add

    ' Copiar los valores a las columnas correspondientes
    nuevaFila.Range(1, 1).Value = Year(Date)
    nuevaFila.Range(1, 3).Value = wsCarlosCobo.Range("C1").Value ' Columna 2
    nuevaFila.Range(1, 4).Value = wsCarlosCobo.Range("A2").Value ' Columna 3
    nuevaFila.Range(1, 5).Value = wsEstadoSem.Range("M4").Value ' Columna 4
    nuevaFila.Range(1, 6).Value = wsEstadoSem.Range("M5").Value ' Columna 5
    nuevaFila.Range(1, 10).Value = Now

    ' Calcular el nombre del mes correspondiente a la semana en la columna B
    nuevaFila.Range(1, 2).Value = ObtenerNombreMesPorSemana(nuevaFila.Range(1, 3).Value)

    MsgBox "Datos importados a la tabla mensual correctamente.", vbInformation

End Sub

Function ObtenerNombreMesPorSemana(numSemana As Variant) As String
    Dim mesNombre As String
    
    ' Asegúrate de que el número de semana sea un valor numérico
    If IsNumeric(numSemana) And numSemana >= 1 And numSemana <= 52 Then
        ' Determinar el mes correspondiente basándose en el número de semana
        Select Case numSemana
            Case 1 To 4: mesNombre = "Enero"
            Case 5 To 8: mesNombre = "Febrero"
            Case 9 To 13: mesNombre = "Marzo"
            Case 14 To 17: mesNombre = "Abril"
            Case 18 To 22: mesNombre = "Mayo"
            Case 23 To 26: mesNombre = "Junio"
            Case 27 To 30: mesNombre = "Julio"
            Case 31 To 35: mesNombre = "Agosto"
            Case 36 To 39: mesNombre = "Septiembre"
            Case 40 To 44: mesNombre = "Octubre"
            Case 45 To 48: mesNombre = "Noviembre"
            Case 49 To 52: mesNombre = "Diciembre"
            Case Else: mesNombre = "Mes Inválido" ' Manejo de errores
        End Select
    Else
        mesNombre = "Semana Inválida" ' Manejo de errores
    End If

    ObtenerNombreMesPorSemana = mesNombre
End Function

Sub CalcularNumeroSemanaYMes()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim textoRango As String
    Dim fechaInicio As Date
    Dim numeroSemana As Long
    Dim nombreMes As String
    Dim i As Long

    ' Establecer la hoja y tabla
    Set ws = ThisWorkbook.Sheets("% Mensual")
    Set tbl = ws.ListObjects("Mensual")

    ' Iterar sobre cada fila de la tabla
    For i = 1 To tbl.ListRows.count
        ' Obtener el valor de la columna 4 (el rango de fechas en formato "20-11 AL 26-11")
        textoRango = tbl.DataBodyRange(i, 4).Value
        
        ' Extraer la fecha de inicio (primer conjunto de día-mes)
        If Len(textoRango) >= 5 Then
            ' Convertir el texto de la primera fecha en un formato de fecha válida
            fechaInicio = DateValue(Mid(textoRango, 1, 2) & "-" & Mid(textoRango, 4, 2) & "-" & Year(Date))
            
            ' Calcular el número de semana de la fecha de inicio
            numeroSemana = WorksheetFunction.WeekNum(fechaInicio, 2)
            
            ' Obtener el nombre del mes correspondiente a la fecha de inicio
            nombreMes = Format(fechaInicio, "mmmm")
            
            ' Colocar el nombre del mes en la columna 2
            tbl.DataBodyRange(i, 2).Value = nombreMes
            
            ' Colocar el número de semana en la columna 3
            tbl.DataBodyRange(i, 3).Value = numeroSemana
        Else
            ' Si el formato de la fecha no es correcto, dejar las celdas vacías o con mensaje
            tbl.DataBodyRange(i, 2).Value = "Formato Inválido"
            tbl.DataBodyRange(i, 3).Value = "Formato Inválido"
        End If
    Next i
End Sub

Sub CalcularEstadisticasAnuales()
    Dim ws As Worksheet
    Dim tabla As ListObject
    Dim yearToFilter As Long
    Dim total As Double
    Dim count As Long
    Dim sumSquares As Double
    Dim promedio As Double
    Dim desvEst As Double
    Dim totalI As Double
    Dim promedioI As Double
    Dim i As Long
    Dim currentYear As Long

    ' Definir la hoja y la tabla
    Set ws = ThisWorkbook.Sheets("% Mensual")
    Set tabla = ws.ListObjects("Mensual")

    ' Obtener el año de la celda L2
    yearToFilter = ws.Range("L3").Value

    ' Inicializar variables
    total = 0
    count = 0
    sumSquares = 0
    totalI = 0

    ' Recorrer las filas de la tabla
    For i = 1 To tabla.ListRows.count
        currentYear = tabla.ListColumns(1).DataBodyRange.Cells(i, 1).Value
        
        ' Verificar si el año de la columna 1 coincide con el año ingresado
        If currentYear = yearToFilter Then
            ' Verificar si el valor de la columna 7 ("% Cobrado") es numérico
            If IsNumeric(tabla.ListColumns(8).DataBodyRange.Cells(i, 1).Value) Then
                total = total + tabla.ListColumns(8).DataBodyRange.Cells(i, 1).Value
                sumSquares = sumSquares + tabla.ListColumns(8).DataBodyRange.Cells(i, 1).Value ^ 2
                count = count + 1
            End If
            
            ' Verificar si el valor de la columna I es numérico
            If IsNumeric(tabla.ListColumns(9).DataBodyRange.Cells(i, 1).Value) Then
                totalI = totalI + tabla.ListColumns(9).DataBodyRange.Cells(i, 1).Value
            End If
        End If
    Next i

    ' Calcular promedio y desviación estándar para la columna 7
    If count > 0 Then
        promedio = total / count
        desvEst = Sqr((sumSquares / count) - (promedio ^ 2))
    Else
        promedio = 0
        desvEst = 0
    End If

    ' Calcular promedio para la columna I
    If count > 0 Then
        promedioI = totalI / count
    Else
        promedioI = 0
    End If

    ' Poner los resultados en las celdas M2, N2 y O2
    ws.Range("M3").Value = promedio
    ws.Range("N3").Value = desvEst
    ws.Range("O3").Value = promedioI
    
    Debug.Print "Promedio: " & promedio & ", Desviación Estándar: " & desvEst & ", Promedio I: " & promedioI
End Sub










