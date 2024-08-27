Attribute VB_Name = "FiltradoDatos2"
Sub FiltrarDatos2()
    Dim wsEvaluacion As Worksheet
    Dim wsResultados As Worksheet
    Dim tblEvaluacion As ListObject
    Dim evalRow As ListRow
    Dim data80Plus As Collection
    Dim data60To80 As Collection
    Dim data60Minus As Collection
    Dim i As Long
    
    Set wsEvaluacion = ThisWorkbook.Sheets("Evaluacion")
    Set wsResultados = ThisWorkbook.Sheets("Resultados semestre")
    Set tblEvaluacion = wsEvaluacion.ListObjects("Tabla1")
    
    wsResultados.Range("G5:L" & wsResultados.Rows.count).ClearContents
    
    ' Colecciones para almacenar los datos
    Set data80Plus = New Collection
    Set data60To80 = New Collection
    Set data60Minus = New Collection
    
    ' Filtrar y almacenar datos en colecciones
    For Each evalRow In tblEvaluacion.ListRows
        If evalRow.Range(28).Value >= 32 Then
            data80Plus.Add Array(CStr(evalRow.Range(2).Value) & CStr(evalRow.Range(3).Value), evalRow.Range(28).Value)
        ElseIf evalRow.Range(28).Value > 24 And evalRow.Range(28).Value < 32 Then
            data60To80.Add Array(CStr(evalRow.Range(2).Value) & CStr(evalRow.Range(3).Value), evalRow.Range(28).Value)
        ElseIf evalRow.Range(28).Value <= 24 Then
            data60Minus.Add Array(CStr(evalRow.Range(2).Value) & CStr(evalRow.Range(3).Value), evalRow.Range(28).Value)
        End If
    Next evalRow
    
    ' Ordenar colecciones de mayor a menor
    Call SortCollection(data80Plus)
    Call SortCollection(data60To80)
    Call SortCollection(data60Minus)
    
    ' Transferir datos ordenados a la hoja de resultados
    For i = 1 To data80Plus.count
        wsResultados.Cells(5 + i - 1, 7).Value = data80Plus(i)(0)
        wsResultados.Cells(5 + i - 1, 8).Value = data80Plus(i)(1)
    Next i
    
    For i = 1 To data60To80.count
        wsResultados.Cells(5 + i - 1, 9).Value = data60To80(i)(0)
        wsResultados.Cells(5 + i - 1, 10).Value = data60To80(i)(1)
    Next i
    
    For i = 1 To data60Minus.count
        wsResultados.Cells(5 + i - 1, 11).Value = data60Minus(i)(0)
        wsResultados.Cells(5 + i - 1, 12).Value = data60Minus(i)(1)
    Next i
    
    Set wsEvaluacion = Nothing
    Set wsResultados = Nothing
    Set tblEvaluacion = Nothing
    Set data80Plus = Nothing
    Set data60To80 = Nothing
    Set data60Minus = Nothing
End Sub

Sub SortCollection(coll As Collection)
    Dim i As Long, j As Long
    Dim tempArray() As Variant
    Dim temp As Variant
    
    ' Verificar si la colección está vacía
    If coll.count = 0 Then Exit Sub
    
    ' Convertir la colección a una matriz
    ReDim tempArray(1 To coll.count)
    For i = 1 To coll.count
        tempArray(i) = coll(i)
    Next i
    
    ' Ordenar la matriz
    For i = 1 To UBound(tempArray) - 1
        For j = i + 1 To UBound(tempArray)
            If tempArray(i)(1) < tempArray(j)(1) Then
                temp = tempArray(i)
                tempArray(i) = tempArray(j)
                tempArray(j) = temp
            End If
        Next j
    Next i
    
    ' Limpiar la colección original
    Do While coll.count > 0
        coll.Remove 1
    Loop
    
    ' Volver a llenar la colección con los datos ordenados
    For i = 1 To UBound(tempArray)
        coll.Add tempArray(i)
    Next i
End Sub



