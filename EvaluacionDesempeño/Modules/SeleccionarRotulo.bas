Attribute VB_Name = "SeleccionarRotulo"
' BOTON PARA SELECCIONAR EL INFORME GENERADO AUTOMATICAMENTE
Sub SeleccionarRango()
    Dim wsRotulo As Worksheet
    Dim UltimaFila As Long
    
    ' Hoja de cálculo a tomar
    Set wsRotulo = Worksheets("ROTULO")
    
    ' Encuentra la última fila ocupada en la columna A
    UltimaFila = wsRotulo.Cells(Rows.count, "A").End(xlUp).Row
    
    ' Seleccionar el rango desde A1 hasta la última fila del rótulo
    wsRotulo.Range("A1:F" & UltimaFila).Select
End Sub

