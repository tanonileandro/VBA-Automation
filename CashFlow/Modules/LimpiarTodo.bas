Attribute VB_Name = "LimpiarTodo"
Sub LimpiarTodo()
    
    ResetColumnsToZero
    ResetColumnsToZeroB
    
End Sub

Sub ResetColumnsToZero()
    Dim wsCarteraPagos As Worksheet
    Dim i As Long
    
    Set wsCarteraPagos = ThisWorkbook.Sheets("CARTERA-PAGOS")
    
    For i = 3 To 69
        wsCarteraPagos.Range("E" & i & ":F" & i).Value = 0
    Next i
End Sub

Sub ResetColumnsToZeroB()
    Dim wsCarteraPagos As Worksheet
    Dim i As Long
    
    Set wsCarteraPagos = ThisWorkbook.Sheets("CARTERA-PAGOS")
    
    For i = 3 To 69
        wsCarteraPagos.Range("J" & i & ":K" & i).Value = 0
    Next i
End Sub



