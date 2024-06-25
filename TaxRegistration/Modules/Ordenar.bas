Attribute VB_Name = "Ordenar"
Sub OrdenarMeses_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim sortRange As Range

        ' Configurar la hoja activa y la tabla
    Set ws = ActiveSheet
    On Error Resume Next
    Set tbl = ws.ListObjects(1)  ' Intenta obtener la primera tabla en la hoja activa
    On Error GoTo 0

    ' Verificar que la tabla y la hoja son correctas
    If tbl Is Nothing Then
        MsgBox "No se encontró ninguna tabla en la hoja activa.", vbCritical
        Exit Sub
    End If
    
    ' Verifica que la tabla y la hoja son correctas
    If ws Is Nothing Then
        MsgBox "La hoja 'ImpAnual' no se encontró.", vbCritical
        Exit Sub
    End If
    If tbl Is Nothing Then
        MsgBox "La tabla 'Tabla3' no se encontró en la hoja 'ImpAnual'.", vbCritical
        Exit Sub
    End If
    
    ' Definir el rango a ordenar (columna A)
    Set sortRange = tbl.ListColumns("Mes").DataBodyRange  ' Ajusta "Mes" al nombre de la columna donde están los meses
    
    ' Ordenar los meses de enero a diciembre
    sortRange.Sort Key1:=sortRange, Order1:=xlAscending, Header:=xlYes
    
    MsgBox "Los meses se han ordenado correctamente.", vbInformation
End Sub

