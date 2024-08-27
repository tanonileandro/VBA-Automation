Attribute VB_Name = "Exportar"
Sub ExportarTodos()
    ' Limpiar todas las tablas antes de exportar
    LimpiarTablas
    
    ' Proceder con la exportación
    ExportarCarlosCobo
    ExportarDiegoPicci
    ExportarHoracioSchaad
    ExportarMarcosNadin
    ExportarPedroIuorno
    ExportarRosarioPack
    ExportarEmbalajes
End Sub

Sub LimpiarTablas()
    LimpiarTabla "Carlos Cobo", "TablaCC"
    LimpiarTabla "Diego Picci", "TablaDP"
    LimpiarTabla "Horacio Schaad", "TablaHS"
    LimpiarTabla "Marcos Nadin", "TablaMN"
    LimpiarTabla "Pedro Iuorno", "TablaPI"
    LimpiarTabla "Rosario Pack", "TablaRP"
    LimpiarTabla "Embalajes", "TablaE"
End Sub

Sub LimpiarTabla(sheetName As String, tableName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        On Error Resume Next
        Set tbl = ws.ListObjects(tableName)
        On Error GoTo 0
        
        If Not tbl Is Nothing Then
            ' Eliminar todas las filas de la tabla
            If tbl.ListRows.Count > 0 Then
                tbl.DataBodyRange.Delete
            End If
        Else
            MsgBox "No se encontró la tabla " & tableName & " en la hoja " & sheetName, vbExclamation
        End If
    Else
        MsgBox "No se encontró la hoja " & sheetName, vbExclamation
    End If
End Sub
Sub ExportarCarlosCobo()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsProveedores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaProveedores As Long
    Dim tblDestino As ListObject
    Dim i As Long
    Dim columnasAColar As Variant
    Dim j As Long
    Dim valor As Variant

    Set wsProveedores = ThisWorkbook.Sheets("COBRANZA TOTAL")

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Carlos Cobo")
    On Error GoTo 0

    ultimaFilaProveedores = wsProveedores.Cells(wsProveedores.Rows.Count, "C").End(xlUp).Row

    ' Verificar la presencia de valores N/A en la columna de interés
    hayNAs = False
    For i = 5 To ultimaFilaProveedores
        If IsError(wsProveedores.Cells(i, 5).Value) Then
            If wsProveedores.Cells(i, 5).Value = CVErr(xlErrNA) Then
                hayNAs = True
                Exit For
            End If
        End If
    Next i

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("TablaCC")
    On Error GoTo 0
    
    If tblDestino Is Nothing Then
        MsgBox "No se encontró la tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    ' Columnas que queremos copiar (en el orden deseado)
    columnasAColar = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15)

    For i = 5 To ultimaFilaProveedores
        If Not IsEmpty(wsProveedores.Cells(i, 5).Value) Then
            ' Compara los valores de la columna C después de aplicar Trim y UCase
            If UCase(Trim(wsProveedores.Cells(i, 3).Value)) = "CARLOS COBO" Then
                tblDestino.ListRows.Add
                ' Copiar solo las columnas específicas
                For j = LBound(columnasAColar) To UBound(columnasAColar)
                    valor = wsProveedores.Cells(i, columnasAColar(j)).Value
                    tblDestino.ListRows(tblDestino.ListRows.Count).Range.Cells(1, j + 1).Value = valor
                Next j
            End If
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ExportarDiegoPicci()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsProveedores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaProveedores As Long
    Dim tblDestino As ListObject
    Dim i As Long
    Dim columnasAColar As Variant
    Dim j As Long
    Dim valor As Variant

    Set wsProveedores = ThisWorkbook.Sheets("COBRANZA TOTAL")

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Diego Picci")
    On Error GoTo 0

    ultimaFilaProveedores = wsProveedores.Cells(wsProveedores.Rows.Count, "C").End(xlUp).Row

    ' Verificar la presencia de valores N/A en la columna de interés
    hayNAs = False
    For i = 5 To ultimaFilaProveedores
        If IsError(wsProveedores.Cells(i, 5).Value) Then
            If wsProveedores.Cells(i, 5).Value = CVErr(xlErrNA) Then
                hayNAs = True
                Exit For
            End If
        End If
    Next i

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("TablaDP")
    On Error GoTo 0
    
    If tblDestino Is Nothing Then
        MsgBox "No se encontró la tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    ' Columnas que queremos copiar (en el orden deseado)
    columnasAColar = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15)

    For i = 5 To ultimaFilaProveedores
        If Not IsEmpty(wsProveedores.Cells(i, 5).Value) Then
            ' Compara los valores de la columna C después de aplicar Trim y UCase
            If UCase(Trim(wsProveedores.Cells(i, 3).Value)) = "DIEGO PICCI" Then
                tblDestino.ListRows.Add
                ' Copiar solo las columnas específicas
                For j = LBound(columnasAColar) To UBound(columnasAColar)
                    valor = wsProveedores.Cells(i, columnasAColar(j)).Value
                    tblDestino.ListRows(tblDestino.ListRows.Count).Range.Cells(1, j + 1).Value = valor
                Next j
            End If
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ExportarHoracioSchaad()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsProveedores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaProveedores As Long
    Dim tblDestino As ListObject
    Dim i As Long
    Dim columnasAColar As Variant
    Dim j As Long
    Dim valor As Variant

    Set wsProveedores = ThisWorkbook.Sheets("COBRANZA TOTAL")

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Horacio Schaad")
    On Error GoTo 0

    ultimaFilaProveedores = wsProveedores.Cells(wsProveedores.Rows.Count, "C").End(xlUp).Row

    ' Verificar la presencia de valores N/A en la columna de interés
    hayNAs = False
    For i = 5 To ultimaFilaProveedores
        If IsError(wsProveedores.Cells(i, 5).Value) Then
            If wsProveedores.Cells(i, 5).Value = CVErr(xlErrNA) Then
                hayNAs = True
                Exit For
            End If
        End If
    Next i

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("TablaHS")
    On Error GoTo 0
    
    If tblDestino Is Nothing Then
        MsgBox "No se encontró la tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    ' Columnas que queremos copiar (en el orden deseado)
    columnasAColar = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15)

    For i = 5 To ultimaFilaProveedores
        If Not IsEmpty(wsProveedores.Cells(i, 5).Value) Then
            ' Compara los valores de la columna C después de aplicar Trim y UCase
            If UCase(Trim(wsProveedores.Cells(i, 3).Value)) = "HORACIO SCHAAD" Then
                tblDestino.ListRows.Add
                ' Copiar solo las columnas específicas
                For j = LBound(columnasAColar) To UBound(columnasAColar)
                    valor = wsProveedores.Cells(i, columnasAColar(j)).Value
                    tblDestino.ListRows(tblDestino.ListRows.Count).Range.Cells(1, j + 1).Value = valor
                Next j
            End If
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ExportarMarcosNadin()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsProveedores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaProveedores As Long
    Dim tblDestino As ListObject
    Dim i As Long
    Dim columnasAColar As Variant
    Dim j As Long
    Dim valor As Variant

    Set wsProveedores = ThisWorkbook.Sheets("COBRANZA TOTAL")

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Marcos Nadin")
    On Error GoTo 0

    ultimaFilaProveedores = wsProveedores.Cells(wsProveedores.Rows.Count, "C").End(xlUp).Row

    ' Verificar la presencia de valores N/A en la columna de interés
    hayNAs = False
    For i = 5 To ultimaFilaProveedores
        If IsError(wsProveedores.Cells(i, 5).Value) Then
            If wsProveedores.Cells(i, 5).Value = CVErr(xlErrNA) Then
                hayNAs = True
                Exit For
            End If
        End If
    Next i

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("TablaMN")
    On Error GoTo 0
    
    If tblDestino Is Nothing Then
        MsgBox "No se encontró la tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    ' Columnas que queremos copiar (en el orden deseado)
    columnasAColar = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15)

    For i = 5 To ultimaFilaProveedores
        If Not IsEmpty(wsProveedores.Cells(i, 5).Value) Then
            ' Compara los valores de la columna C después de aplicar Trim y UCase
            If UCase(Trim(wsProveedores.Cells(i, 3).Value)) = "MARCOS NADIN" Then
                tblDestino.ListRows.Add
                ' Copiar solo las columnas específicas
                For j = LBound(columnasAColar) To UBound(columnasAColar)
                    valor = wsProveedores.Cells(i, columnasAColar(j)).Value
                    tblDestino.ListRows(tblDestino.ListRows.Count).Range.Cells(1, j + 1).Value = valor
                Next j
            End If
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ExportarPedroIuorno()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsProveedores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaProveedores As Long
    Dim tblDestino As ListObject
    Dim i As Long
    Dim columnasAColar As Variant
    Dim j As Long
    Dim valor As Variant

    Set wsProveedores = ThisWorkbook.Sheets("COBRANZA TOTAL")

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Pedro Iuorno")
    On Error GoTo 0

    ultimaFilaProveedores = wsProveedores.Cells(wsProveedores.Rows.Count, "C").End(xlUp).Row

    ' Verificar la presencia de valores N/A en la columna de interés
    hayNAs = False
    For i = 5 To ultimaFilaProveedores
        If IsError(wsProveedores.Cells(i, 5).Value) Then
            If wsProveedores.Cells(i, 5).Value = CVErr(xlErrNA) Then
                hayNAs = True
                Exit For
            End If
        End If
    Next i

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("TablaPI")
    On Error GoTo 0
    
    If tblDestino Is Nothing Then
        MsgBox "No se encontró la tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    ' Columnas que queremos copiar (en el orden deseado)
    columnasAColar = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15)

    For i = 5 To ultimaFilaProveedores
        If Not IsEmpty(wsProveedores.Cells(i, 5).Value) Then
            ' Compara los valores de la columna C después de aplicar Trim y UCase
            If UCase(Trim(wsProveedores.Cells(i, 3).Value)) = "PEDRO IUORNO" Then
                tblDestino.ListRows.Add
                ' Copiar solo las columnas específicas
                For j = LBound(columnasAColar) To UBound(columnasAColar)
                    valor = wsProveedores.Cells(i, columnasAColar(j)).Value
                    tblDestino.ListRows(tblDestino.ListRows.Count).Range.Cells(1, j + 1).Value = valor
                Next j
            End If
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ExportarRosarioPack()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsProveedores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaProveedores As Long
    Dim tblDestino As ListObject
    Dim i As Long
    Dim columnasAColar As Variant
    Dim j As Long
    Dim valor As Variant

    Set wsProveedores = ThisWorkbook.Sheets("COBRANZA TOTAL")

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Rosario Pack")
    On Error GoTo 0

    ultimaFilaProveedores = wsProveedores.Cells(wsProveedores.Rows.Count, "C").End(xlUp).Row

    ' Verificar la presencia de valores N/A en la columna de interés
    hayNAs = False
    For i = 5 To ultimaFilaProveedores
        If IsError(wsProveedores.Cells(i, 5).Value) Then
            If wsProveedores.Cells(i, 5).Value = CVErr(xlErrNA) Then
                hayNAs = True
                Exit For
            End If
        End If
    Next i

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("TablaRP")
    On Error GoTo 0
    
    If tblDestino Is Nothing Then
        MsgBox "No se encontró la tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    ' Columnas que queremos copiar (en el orden deseado)
    columnasAColar = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15)

    For i = 5 To ultimaFilaProveedores
        If Not IsEmpty(wsProveedores.Cells(i, 5).Value) Then
            ' Compara los valores de la columna C después de aplicar Trim y UCase
            If UCase(Trim(wsProveedores.Cells(i, 3).Value)) = "ROSARIO PACK" Then
                tblDestino.ListRows.Add
                ' Copiar solo las columnas específicas
                For j = LBound(columnasAColar) To UBound(columnasAColar)
                    valor = wsProveedores.Cells(i, columnasAColar(j)).Value
                    tblDestino.ListRows(tblDestino.ListRows.Count).Range.Cells(1, j + 1).Value = valor
                Next j
            End If
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
Sub ExportarEmbalajes()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsProveedores As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaProveedores As Long
    Dim tblDestino As ListObject
    Dim i As Long
    Dim columnasAColar As Variant
    Dim j As Long
    Dim valor As Variant

    Set wsProveedores = ThisWorkbook.Sheets("COBRANZA TOTAL")

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Embalajes")
    On Error GoTo 0

    ultimaFilaProveedores = wsProveedores.Cells(wsProveedores.Rows.Count, "C").End(xlUp).Row

    ' Verificar la presencia de valores N/A en la columna de interés
    hayNAs = False
    For i = 5 To ultimaFilaProveedores
        If IsError(wsProveedores.Cells(i, 5).Value) Then
            If wsProveedores.Cells(i, 5).Value = CVErr(xlErrNA) Then
                hayNAs = True
                Exit For
            End If
        End If
    Next i

    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("TablaE")
    On Error GoTo 0
    
    If tblDestino Is Nothing Then
        MsgBox "No se encontró la tabla en la hoja de destino.", vbExclamation
        Exit Sub
    End If

    ' Columnas que queremos copiar (en el orden deseado)
    columnasAColar = Array(1, 2, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15)

    For i = 5 To ultimaFilaProveedores
        If Not IsEmpty(wsProveedores.Cells(i, 5).Value) Then
            ' Compara los valores de la columna C después de aplicar Trim y UCase
            If UCase(Trim(wsProveedores.Cells(i, 3).Value)) = "EMBALAJES" Then
                tblDestino.ListRows.Add
                ' Copiar solo las columnas específicas
                For j = LBound(columnasAColar) To UBound(columnasAColar)
                    valor = wsProveedores.Cells(i, columnasAColar(j)).Value
                    tblDestino.ListRows(tblDestino.ListRows.Count).Range.Cells(1, j + 1).Value = valor
                Next j
            End If
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

