Attribute VB_Name = "SeleccionarRotulo"
' BOTON PARA SELECCIONAR EL INFORME GENERADO AUTOMATICAMENTE
Sub SeleccionarRango()
    Dim wsRotulo As Worksheet
    Dim UltimaFila As Long
    
    ' Hoja de c�lculo a tomar
    Set wsRotulo = Worksheets("ROTULO")
    
    ' Encuentra la �ltima fila ocupada en la columna A
    UltimaFila = wsRotulo.Cells(Rows.count, "A").End(xlUp).Row
    
    ' Seleccionar el rango desde A1 hasta la �ltima fila del r�tulo
    wsRotulo.Range("A1:F" & UltimaFila).Select
End Sub

