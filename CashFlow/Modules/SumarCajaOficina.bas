Attribute VB_Name = "SumarCajaOficina"
Sub SumarCajaOficina()
    Dim wsCartera As Worksheet
    Dim wsMensual As Worksheet
    Dim lastRow As Long
    Dim sumaCajaOficina As Double
    Dim sumaDemo As Double
    Dim i As Long
    
    ' Define las hojas de trabajo
    Set wsCartera = ThisWorkbook.Sheets("Cartera Chq")
    Set wsMensual = ThisWorkbook.Sheets("Mensual")
    
    ' Encuentra la última fila con datos en la hoja Cartera Chq
    lastRow = wsCartera.Cells(wsCartera.Rows.Count, "E").End(xlUp).Row
    
    ' Inicializa las sumas
    sumaCajaOficina = 0
    sumaDemo = 0
    
    ' Itera a través de los datos en la hoja Cartera Chq
    For i = 2 To lastRow ' Empieza en 2 para omitir el encabezado
        If wsCartera.Cells(i, "E").Value = "Caja Oficina" Then
            ' Suma el valor de la columna I si en la columna E dice "Caja Oficina"
            sumaCajaOficina = sumaCajaOficina + wsCartera.Cells(i, "I").Value
        ElseIf wsCartera.Cells(i, "E").Value = "Demo" Then
            ' Suma el valor de la columna I si en la columna E dice "Demo"
            sumaDemo = sumaDemo + wsCartera.Cells(i, "I").Value
        End If
    Next i
    
    ' Pega el resultado en la celda E17 de la hoja Mensual
    wsMensual.Range("E17").Value = sumaCajaOficina
    
    ' Pega el resultado en la celda F17 de la hoja Mensual
    wsMensual.Range("F17").Value = sumaDemo
End Sub


