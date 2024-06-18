Attribute VB_Name = "SemanalVerificarCajasAB"
Sub SemanalVerificarCajasAB()
    VerificarCajaOficial
    VerificarDemo
End Sub

Sub VerificarCajaOficial()
    Dim wsMensual As Worksheet
    Dim wsCarteraChq As Worksheet
    Dim valorBuscadoD13 As Variant
    Dim valorBuscadoD27 As Variant
    Dim valorBuscadoD34 As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim sumaD13 As Double
    Dim sumaD27 As Double
    Dim sumaD34 As Double
    
    Set wsMensual = ThisWorkbook.Sheets("Semanal")
    Set wsCarteraChq = ThisWorkbook.Sheets("Cartera Chq")
    
    ' Valor original de la celda A13
    valorBuscadoD13 = wsMensual.Range("A13").Value
    
    ' Valores incrementados para la búsqueda en D27 y D34
    valorBuscadoD27 = valorBuscadoD13 + 1
    valorBuscadoD34 = valorBuscadoD13 + 2
    
    sumaD13 = 0
    sumaD27 = 0
    sumaD34 = 0
    
    lastRow = wsCarteraChq.Cells(wsCarteraChq.Rows.Count, "A").End(xlUp).Row
    
    ' Lógica para D13
    For i = 1 To lastRow
        If wsCarteraChq.Cells(i, 1).Value <= valorBuscadoD13 Then
            If wsCarteraChq.Cells(i, 5).Value = "Caja Oficina" Then
                sumaD13 = sumaD13 + wsCarteraChq.Cells(i, 9).Value
            End If
        End If
    Next i
    
    ' Lógica para D27
    For i = 1 To lastRow
        If wsCarteraChq.Cells(i, 1).Value = valorBuscadoD27 Then
            If wsCarteraChq.Cells(i, 5).Value = "Caja Oficina" Then
                sumaD27 = sumaD27 + wsCarteraChq.Cells(i, 9).Value
            End If
        End If
    Next i
    
    ' Lógica para D34
    For i = 1 To lastRow
        If wsCarteraChq.Cells(i, 1).Value = valorBuscadoD34 Then
            If wsCarteraChq.Cells(i, 5).Value = "Caja Oficina" Then
                sumaD34 = sumaD34 + wsCarteraChq.Cells(i, 9).Value
            End If
        End If
    Next i
    
    ' Escribir las sumas en las celdas D13, D27 y D34
    wsMensual.Range("D13").Value = sumaD13
    wsMensual.Range("D27").Value = sumaD27
    wsMensual.Range("D34").Value = sumaD34

End Sub

Sub VerificarDemo()
    Dim wsMensual As Worksheet
    Dim wsCarteraChq As Worksheet
    Dim valorBuscadoD13 As Variant
    Dim valorBuscadoD27 As Variant
    Dim valorBuscadoD34 As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim sumaD13 As Double
    Dim sumaD27 As Double
    Dim sumaD34 As Double
    
    Set wsMensual = ThisWorkbook.Sheets("Semanal")
    Set wsCarteraChq = ThisWorkbook.Sheets("Cartera Chq")
    
    ' Valor original de la celda A13
    valorBuscadoD13 = wsMensual.Range("A13").Value
    
    ' Valores incrementados para la búsqueda en D27 y D34
    valorBuscadoD27 = valorBuscadoD13 + 1
    valorBuscadoD34 = valorBuscadoD13 + 2
    
    sumaD13 = 0
    sumaD27 = 0
    sumaD34 = 0
    
    lastRow = wsCarteraChq.Cells(wsCarteraChq.Rows.Count, "A").End(xlUp).Row
    
    ' Lógica para D13
    For i = 1 To lastRow
        If wsCarteraChq.Cells(i, 1).Value <= valorBuscadoD13 Then
            If wsCarteraChq.Cells(i, 5).Value = "Demo" Then
                sumaD13 = sumaD13 + wsCarteraChq.Cells(i, 9).Value
            End If
        End If
    Next i
    
    ' Lógica para D27
    For i = 1 To lastRow
        If wsCarteraChq.Cells(i, 1).Value = valorBuscadoD27 Then
            If wsCarteraChq.Cells(i, 5).Value = "Demo" Then
                sumaD27 = sumaD27 + wsCarteraChq.Cells(i, 9).Value
            End If
        End If
    Next i
    
    ' Lógica para D34
    For i = 1 To lastRow
        If wsCarteraChq.Cells(i, 1).Value = valorBuscadoD34 Then
            If wsCarteraChq.Cells(i, 5).Value = "Demo" Then
                sumaD34 = sumaD34 + wsCarteraChq.Cells(i, 9).Value
            End If
        End If
    Next i
    
    ' Escribir las sumas en las celdas D13, D27 y D34
    wsMensual.Range("E13").Value = sumaD13
    wsMensual.Range("E27").Value = sumaD27
    wsMensual.Range("E34").Value = sumaD34

End Sub
