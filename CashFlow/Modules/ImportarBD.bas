Attribute VB_Name = "ImportarBD"
Sub ImportarBD()
    LimpiarBD
    ImportarPagoProveedores
End Sub

Sub LimpiarBD()
    On Error Resume Next
    
    Dim wsDestino As Worksheet
    Dim tblDestino As ListObject

    ' Desactivar actualizaciones de pantalla y cálculos
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Definir la hoja de destino
    Set wsDestino = ThisWorkbook.Sheets("CARTERA-PAGOS")
    
    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("Tabla2")
    On Error GoTo 0

    If tblDestino Is Nothing Then
        MsgBox "La tabla de destino 'Tabla2' no se encontró en la hoja 'CARTERA-PAGOS'.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si la tabla tiene datos
    If tblDestino.ListRows.Count = 0 Then
        Exit Sub
    End If

    ' Limpiar la tabla eliminando todas las filas
    tblDestino.DataBodyRange.Delete
    
    ' Reactivar actualizaciones de pantalla y cálculos
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub


Sub ImportarPagoProveedores()
    On Error Resume Next
    
    Dim wbOrigen As Workbook
    Dim wsOrigen1 As Worksheet
    Dim wsOrigen2 As Worksheet
    Dim wsOrigen3 As Worksheet
    Dim wsOrigen4 As Worksheet
    Dim wsDestino As Worksheet
    Dim tblDestino As ListObject
    Dim lastRowDestino As Long
    Dim i As Long
    Dim j As Long
    Dim visibleRow As Boolean

    ' Desactivar actualizaciones de pantalla y cálculos
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Abrir el libro de origen sin mostrarlo
    Set wbOrigen = Workbooks.Open(Filename:="Y:\PROVEEDORES\PAGO A PROVEEDORES\Planilla_Pagos_2024.xlsm", UpdateLinks:=False, ReadOnly:=True)
    
    ' Definir las hojas de origen
    Set wsOrigen1 = wbOrigen.Sheets("CHEQUES A")
    Set wsOrigen2 = wbOrigen.Sheets("PAPELERA A")
    Set wsOrigen3 = wbOrigen.Sheets("B")
    Set wsOrigen4 = wbOrigen.Sheets("PAPELERA B")
    
    ' Definir la hoja de destino
    Set wsDestino = ThisWorkbook.Sheets("CARTERA-PAGOS")
    
    On Error Resume Next
    Set tblDestino = wsDestino.ListObjects("Tabla2")
    On Error GoTo 0

    If tblDestino Is Nothing Then
        MsgBox "La tabla de destino 'Tabla2' no se encontró en la hoja 'CARTERA-PAGOS'.", vbExclamation
        wbOrigen.Close False ' Cerrar el libro de origen si hay un error
        Exit Sub
    End If

    ' Definir la última fila en la tabla de destino
    lastRowDestino = tblDestino.ListRows.Count + 1

    ' Copiar los datos visibles de la Tabla4 de la hoja de origen 1 (CHEQUES A)
    For i = 1 To wsOrigen1.ListObjects("Tabla4").ListRows.Count
        visibleRow = True
        For j = 1 To wsOrigen1.ListObjects("Tabla4").ListColumns.Count
            If Not wsOrigen1.ListObjects("Tabla4").ListColumns(j).Range.Cells(i).EntireRow.Hidden Then
                visibleRow = visibleRow And True
            Else
                visibleRow = False
                Exit For
            End If
        Next j
        If visibleRow Then
            tblDestino.ListRows.Add
            tblDestino.ListRows(lastRowDestino).Range.Value = wsOrigen1.ListObjects("Tabla4").ListRows(i).Range.Value
            lastRowDestino = lastRowDestino + 1
        End If
    Next i

    ' Copiar los datos visibles de la Tabla5 de la hoja de origen 2 (PAPELERA A)
    For i = 1 To wsOrigen2.ListObjects("Tabla5").ListRows.Count
        visibleRow = True
        For j = 1 To wsOrigen2.ListObjects("Tabla5").ListColumns.Count
            If Not wsOrigen2.ListObjects("Tabla5").ListColumns(j).Range.Cells(i).EntireRow.Hidden Then
                visibleRow = visibleRow And True
            Else
                visibleRow = False
                Exit For
            End If
        Next j
        If visibleRow Then
            tblDestino.ListRows.Add
            tblDestino.ListRows(lastRowDestino).Range.Value = wsOrigen2.ListObjects("Tabla5").ListRows(i).Range.Value
            lastRowDestino = lastRowDestino + 1
        End If
    Next i
    
    ' Copiar los datos visibles de la Tabla6 de la hoja de origen 3 (B Tabla3)
    For i = 1 To wsOrigen3.ListObjects("Tabla3").ListRows.Count
        visibleRow = True
        For j = 1 To wsOrigen3.ListObjects("Tabla3").ListColumns.Count
            If Not wsOrigen3.ListObjects("Tabla3").ListColumns(j).Range.Cells(i).EntireRow.Hidden Then
                visibleRow = visibleRow And True
            Else
                visibleRow = False
                Exit For
            End If
        Next j
        If visibleRow Then
            tblDestino.ListRows.Add
            tblDestino.ListRows(lastRowDestino).Range.Value = wsOrigen3.ListObjects("Tabla3").ListRows(i).Range.Value
            lastRowDestino = lastRowDestino + 1
        End If
    Next i
    
    ' Copiar los datos visibles de la Tabla7 de la hoja de origen 4 (PAPELERA B Tabla511)
    For i = 1 To wsOrigen4.ListObjects("Tabla511").ListRows.Count
        visibleRow = True
        For j = 1 To wsOrigen4.ListObjects("Tabla511").ListColumns.Count
            If Not wsOrigen4.ListObjects("Tabla511").ListColumns(j).Range.Cells(i).EntireRow.Hidden Then
                visibleRow = visibleRow And True
            Else
                visibleRow = False
                Exit For
            End If
        Next j
        If visibleRow Then
            tblDestino.ListRows.Add
            tblDestino.ListRows(lastRowDestino).Range.Value = wsOrigen4.ListObjects("Tabla511").ListRows(i).Range.Value
            lastRowDestino = lastRowDestino + 1
        End If
    Next i

    ' Reactivar actualizaciones de pantalla y cálculos
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Cerrar el libro de origen
    wbOrigen.Close False
    
End Sub




