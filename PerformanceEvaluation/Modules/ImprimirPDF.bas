Attribute VB_Name = "ImprimirPDF"
Sub ImprimirAFormatoPDF()
    Dim RangoSeleccionado As Range
    Dim NombreArchivoPDF As String
    Dim NumCotizacion As Variant
    Dim NombreCliente As String
    Dim CeldaNumCotizacion As Range
    Dim ExportError As String
    
    ' Obtener el número de cotización y el nombre del cliente
    NumCotizacion = Worksheets("ROTULO").Range("C9").Value
    NombreCliente = Worksheets("ROTULO").Range("C10").Value
    
    ' Verifica si hay alguna selección
    On Error Resume Next
    Set RangoSeleccionado = Selection
    On Error GoTo 0

    ' Verificar si se ha seleccionado un rango
    If RangoSeleccionado Is Nothing Then
        MsgBox "No se ha seleccionado ningún rango. Por favor selecciona el contenido que deseas imprimir en PDF.", vbExclamation
        Exit Sub
    End If

    ' Buscar la celda que contiene el número de cotización en el rango seleccionado
    Set CeldaNumCotizacion = RangoSeleccionado.Find(NumCotizacion, LookIn:=xlValues)
    
    If CeldaNumCotizacion Is Nothing Then
        MsgBox "No se encontró Información Personal en el rango seleccionado.", vbExclamation
        Exit Sub
    End If
    
    ' Genera el nombre del archivo PDF con el número de cotización y el nombre del cliente
    NombreArchivoPDF = NumCotizacion & "_" & NombreCliente & ".pdf"
    
    ' Pide al usuario el nombre del archivo PDF y la ubicación
    NombreArchivoPDF = Application.GetSaveAsFilename(FileFilter:="Archivos PDF (*.pdf), *.pdf", _
                                                      Title:="Guardar como PDF", _
                                                      InitialFileName:=NombreArchivoPDF)
    
    ' Comprueba si el usuario ha cancelado
    If NombreArchivoPDF = "Falso" Then
        Exit Sub
    End If
    
    ' Intentar imprimir el rango seleccionado en el archivo PDF
    On Error Resume Next
    RangoSeleccionado.ExportAsFixedFormat Type:=xlTypePDF, Filename:=NombreArchivoPDF
    ExportError = Err.Description
    On Error GoTo 0
    
    If ExportError <> "" Then
        MsgBox "No se pudo guardar el archivo PDF. Motivo: " & ExportError, vbExclamation
    Else
        MsgBox "El archivo PDF se ha creado exitosamente.", vbInformation
    End If
End Sub



