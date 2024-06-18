Attribute VB_Name = "CP_ProcesarCOD"
Sub CP_ProcesarCOD()

    ProcesarCO
    ProcesarDemo
    
End Sub

Sub ProcesarCO()
    Dim wsCartera As Worksheet
    Dim wsPagos As Worksheet
    Dim dictSumas As Object
    Dim ultimaFilaCartera As Long
    Dim i As Long
    
    Set wsCartera = ThisWorkbook.Sheets("Cartera Chq")
    Set wsPagos = ThisWorkbook.Sheets("CARTERA-PAGOS")
    
    Set dictSumas = CreateObject("Scripting.Dictionary")
    
    ultimaFilaCartera = wsCartera.Cells(wsCartera.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To ultimaFilaCartera
        If wsCartera.Cells(i, 5).Value = "Caja Oficina" Then
            clave = wsCartera.Cells(i, 1).Value & "_" & wsCartera.Cells(i, 2).Value
            If Not dictSumas.Exists(clave) Then
                dictSumas.Add clave, wsCartera.Cells(i, 9).Value
            Else
                dictSumas(clave) = dictSumas(clave) + wsCartera.Cells(i, 9).Value
            End If
        End If
    Next i
    
    For i = 3 To wsPagos.Cells(wsPagos.Rows.Count, "D").End(xlUp).Row
    claveDestino = wsPagos.Cells(i, 4).Value & "_" & wsPagos.Cells(i, 3).Value
        If dictSumas.Exists(claveDestino) Then
            wsPagos.Cells(i, 5).Value = Abs(dictSumas(claveDestino))
        End If
    Next i
    
    Set dictSumas = Nothing
End Sub

Sub ProcesarDemo()
    Dim wsCartera As Worksheet
    Dim wsPagos As Worksheet
    Dim dictSumas As Object
    Dim ultimaFilaCartera As Long
    Dim i As Long
    
    Set wsCartera = ThisWorkbook.Sheets("Cartera Chq")
    Set wsPagos = ThisWorkbook.Sheets("CARTERA-PAGOS")
    
    Set dictSumas = CreateObject("Scripting.Dictionary")
    
    ultimaFilaCartera = wsCartera.Cells(wsCartera.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To ultimaFilaCartera
        If wsCartera.Cells(i, 5).Value = "Demo" Then
            clave = wsCartera.Cells(i, 1).Value & "_" & wsCartera.Cells(i, 2).Value
            If Not dictSumas.Exists(clave) Then
                dictSumas.Add clave, wsCartera.Cells(i, 9).Value
            Else
                dictSumas(clave) = dictSumas(clave) + wsCartera.Cells(i, 9).Value
            End If
        End If
    Next i
    
    For i = 3 To wsPagos.Cells(wsPagos.Rows.Count, "D").End(xlUp).Row
    claveDestino = wsPagos.Cells(i, 4).Value & "_" & wsPagos.Cells(i, 3).Value
        If dictSumas.Exists(claveDestino) Then
            wsPagos.Cells(i, 6).Value = Abs(dictSumas(claveDestino))
        End If
    Next i
    
    Set dictSumas = Nothing
End Sub
