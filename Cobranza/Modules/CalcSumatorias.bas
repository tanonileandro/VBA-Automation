Attribute VB_Name = "CalcSumatorias"
Sub CalcularSumatorias()

    Dim Hoja As Worksheet
    Dim tabla As ListObject
    Dim sumaColumnaN As Double
    Dim sumaColumnaG_A As Double
    Dim sumaColumnaG_B As Double
    Dim i As Long
    
    Set Hoja = ActiveSheet
    
    On Error Resume Next
    Set tabla = Hoja.ListObjects(1)
    On Error GoTo 0
    
    If tabla Is Nothing Then
        MsgBox "No se encontró ninguna tabla en la hoja activa.", vbExclamation
        Exit Sub
    End If
    
    sumaColumnaN = 0
    sumaColumnaG_A = 0
    sumaColumnaG_B = 0
    
    ' Calcular sumatorias
    For i = 1 To tabla.ListRows.Count
        ' Sumar la columna 14 (Columna N)
        sumaColumnaN = sumaColumnaN + tabla.DataBodyRange(i, 14).Value
        
        ' Sumar la columna 7 (Columna G) donde la columna 4 (Columna D) diga "A"
        If tabla.DataBodyRange(i, 4).Value = "A" Then
            sumaColumnaG_A = sumaColumnaG_A + tabla.DataBodyRange(i, 14).Value
        End If
        
        ' Sumar la columna 7 (Columna G) donde la columna 4 (Columna D) diga "B"
        If tabla.DataBodyRange(i, 4).Value = "B" Then
            sumaColumnaG_B = sumaColumnaG_B + tabla.DataBodyRange(i, 14).Value
        End If
    Next i
    
    ' Pegar las sumatorias en las celdas correspondientes
    Hoja.Range("M3").Value = sumaColumnaN
    Hoja.Range("H1").Value = sumaColumnaG_A
    Hoja.Range("H2").Value = sumaColumnaG_B
    
End Sub

