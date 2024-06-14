Attribute VB_Name = "BorrarPapelera"
Sub BorrarFilasPapeleraNacional()
    On Error Resume Next
    
    Dim wsActiva As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    Set wsActiva = ActiveSheet
    
    On Error Resume Next
    Set tbl = wsActiva.ListObjects(1)
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "No se encontró ninguna tabla en la hoja activa.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = tbl.ListRows.Count To 1 Step -1
        If Trim(UCase(tbl.ListRows(i).Range.Cells(1, "I").Text)) = "PAPELERA NACIONAL" Then
            tbl.ListRows(i).Delete
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Se han borrado las filas de 'PAPELERA NACIONAL' en la hoja activa.", vbInformation
End Sub


