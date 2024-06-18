Attribute VB_Name = "CP_ProcesarAB"
Sub CP_ProcesarAB()
    ProcesarA
    ProcesarB
End Sub

Sub ProcesarA()
    
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim wsPagos As Worksheet
    Dim dictSumas As Object
    Dim ultimaFilaCartera As Long
    Dim i As Long
    
    ' Establecer la hoja "CARTERA-PAGOS" como la hoja de trabajo
    Set wsPagos = ThisWorkbook.Sheets("CARTERA-PAGOS")
    
    Set dictSumas = CreateObject("Scripting.Dictionary")
    
    ' Encontrar la última fila en la hoja "CARTERA-PAGOS"
    ultimaFilaCartera = wsPagos.Cells(wsPagos.Rows.Count, "R").End(xlUp).Row
    
    ' Iterar sobre las filas desde la fila 80 de la hoja "CARTERA-PAGOS"
    For i = 80 To ultimaFilaCartera
        If Trim(wsPagos.Cells(i, 3).Text) = "A" Then ' And UCase(Trim(wsPagos.Cells(i, 11).Text)) = "CHEQUES"
            clave = wsPagos.Cells(i, 18).Value & "_" & wsPagos.Cells(i, 17).Value ' Combina los valores de las columnas 18 y 17 como clave
            If Not dictSumas.Exists(clave) Then
                dictSumas.Add clave, Abs(wsPagos.Cells(i, 5).Value)
            Else
                dictSumas(clave) = dictSumas(clave) + Abs(wsPagos.Cells(i, 5).Value)
            End If
        End If
    Next i
        
    ' Iterar sobre las filas de la hoja "CARTERA-PAGOS" para actualizar los valores en la columna J
    For i = 3 To wsPagos.Cells(wsPagos.Rows.Count, "D").End(xlUp).Row
        claveDestino = wsPagos.Cells(i, 4).Value & "_" & wsPagos.Cells(i, 3).Value ' Combina los valores de las columnas D y C de la hoja "CARTERA-PAGOS" como clave de búsqueda
        If dictSumas.Exists(claveDestino) Then
            wsPagos.Cells(i, 10).Value = Abs(dictSumas(claveDestino))
        End If
    Next i

    Set dictSumas = Nothing

    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
End Sub

Sub ProcesarB()
    
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim wsPagos As Worksheet
    Dim dictSumas As Object
    Dim ultimaFilaCartera As Long
    Dim i As Long
    
    ' Establecer la hoja "CARTERA-PAGOS" como la hoja de trabajo
    Set wsPagos = ThisWorkbook.Sheets("CARTERA-PAGOS")
    
    Set dictSumas = CreateObject("Scripting.Dictionary")
    
    ' Encontrar la última fila en la hoja "CARTERA-PAGOS"
    ultimaFilaCartera = wsPagos.Cells(wsPagos.Rows.Count, "R").End(xlUp).Row
    
    ' Iterar sobre las filas desde la fila 80 de la hoja "CARTERA-PAGOS"
    For i = 80 To ultimaFilaCartera
        If Trim(wsPagos.Cells(i, 3).Text) = "B" And UCase(Trim(wsPagos.Cells(i, 11).Text)) = "CHEQUES" Then
            clave = wsPagos.Cells(i, 18).Value & "_" & wsPagos.Cells(i, 17).Value ' Combina los valores de las columnas 18 y 17 como clave
            If Not dictSumas.Exists(clave) Then
                dictSumas.Add clave, Abs(wsPagos.Cells(i, 5).Value)
            Else
                dictSumas(clave) = dictSumas(clave) + Abs(wsPagos.Cells(i, 5).Value)
            End If
        End If
    Next i
        
    ' Iterar sobre las filas de la hoja "CARTERA-PAGOS" para actualizar los valores en la columna J
    For i = 3 To wsPagos.Cells(wsPagos.Rows.Count, "D").End(xlUp).Row
        claveDestino = wsPagos.Cells(i, 4).Value & "_" & wsPagos.Cells(i, 3).Value ' Combina los valores de las columnas D y C de la hoja "CARTERA-PAGOS" como clave de búsqueda
        If dictSumas.Exists(claveDestino) Then
            wsPagos.Cells(i, 11).Value = Abs(dictSumas(claveDestino))
        End If
    Next i

    Set dictSumas = Nothing

    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
End Sub
