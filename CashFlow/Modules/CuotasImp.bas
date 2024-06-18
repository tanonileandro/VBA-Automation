Attribute VB_Name = "CuotasImp"
Sub LimpiarImportar()

LimpiarDeudaPorProveedor
ImportarDeudaPorProveedor

End Sub

Sub LimpiarDeudaPorProveedor()
    Dim wsDest As Worksheet
    Dim lastRow As Long
    
    Dim nombreHojaDestino As String
    nombreHojaDestino = "Cuotas Importaciones"
    
    ' Asignar la hoja destino
    Set wsDest = ThisWorkbook.Sheets(nombreHojaDestino)
    
    ' Encontrar la última fila con datos en la hoja destino
    lastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    
    ' Verificar si hay datos para borrar
    If lastRow > 1 Then
        ' Borrar los datos y el formato
        wsDest.Rows("2:" & lastRow).ClearFormats
        wsDest.Rows("2:" & lastRow).ClearContents
    Else
        MsgBox "No hay datos para limpiar en la hoja " & nombreHojaDestino & "."
        Exit Sub
    End If
    
    ' Activar la celda A1 después de limpiar los datos
    wsDest.Range("A1").Select
End Sub
Sub ImportarDeudaPorProveedor()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rngSource As Range
    Dim rngDest As Range
    Dim appExcel As Application
    
    ' Deshabilitar actualizaciones de pantalla y eventos
    Set appExcel = Application
    appExcel.ScreenUpdating = False
    appExcel.DisplayAlerts = False
    
    Dim rutaArchivo As String
    rutaArchivo = "Z:\IMPORTACIONES\Importaciones Papel OK v2.0.xlsm"
    
    ' Nombre de la hoja origen y destino
    Dim nombreHojaOrigen As String
    nombreHojaOrigen = "DEUDA POR PROVEEDOR"
    Dim nombreHojaDestino As String
    nombreHojaDestino = "Cuotas Importaciones"
    
    ' Verificar si el archivo está abierto
    On Error Resume Next
    Set wbSource = Workbooks("Importaciones Papel OK v2.0.xlsm")
    On Error GoTo 0
    
    If wbSource Is Nothing Then
        ' Abrir el archivo origen si no está abierto
        On Error Resume Next
        Set wbSource = Workbooks.Open(rutaArchivo, False, True) ' Open read-only, no actualizar enlaces
        On Error GoTo 0
        
        If wbSource Is Nothing Then
            MsgBox "No se pudo abrir el archivo de origen."
            Exit Sub
        End If
    End If
    
    ' Asignar la hoja origen y destino
    Set wsSource = wbSource.Sheets(nombreHojaOrigen)
    Set wsDest = ThisWorkbook.Sheets(nombreHojaDestino)
    
    ' Encontrar la última fila y columna con datos en la hoja origen
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = 27 ' Columna AA (27 en base 1)
    
    ' Definir el rango a copiar (columnas A a AA)
    Set rngSource = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol))
    
    ' Copiar y pegar valores y formato de forma directa y rápida
    rngSource.Copy
    Set rngDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Offset(1) ' Siguiente fila disponible en columna A
    rngDest.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Limpiar portapapeles después de pegar
    Application.CutCopyMode = False
    
    ' Si el archivo origen fue abierto en este proceso, no cerrarlo
    If Not wbSource Is Nothing Then
        If wbSource.Name <> ThisWorkbook.Name Then
            ' Cerrar el libro origen si no es el libro actual
            wbSource.Close SaveChanges:=False
        End If
    End If
    
    ' Restaurar configuración de pantalla y eventos
    appExcel.ScreenUpdating = True
    appExcel.DisplayAlerts = True
    
    MsgBox "Datos copiados exitosamente."
    
End Sub

