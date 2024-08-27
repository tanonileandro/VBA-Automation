VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "btnEliminarFilasVacias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim ultimaFila As Long
    Dim suma As Double
    Dim contenidoA As String
    Dim rngA As Range
    Dim cellA As Range
    Dim cellJ As Range

    ' Establecer la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("PALLETIZADO")

    ' GENERAR FECHA Y HORA AUTOMATICA UNA VEZ QUE SE AGREGA UN CODIGO DE BARRA ----------------------------------------------------------------------------------------
    
    Set rng = Intersect(Target, Me.Range("A4:A" & Me.Rows.Count))
    
    If rng Is Nothing Then Exit Sub

    Application.EnableEvents = False

    For Each cell In rng
        If cell.Value <> "" Then
            cell.Offset(0, 1).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
        Else
            cell.Offset(0, 1).ClearContents
        End If
    Next cell

    Application.EnableEvents = True

    ' VALIDACIÓN DE DUPLICADOS
    If Not Intersect(Target, ws.Range("A4:A" & ws.Rows.Count)) Is Nothing Then
        Application.EnableEvents = False
        For Each cell In Target
            If cell.Value <> "" Then
                If WorksheetFunction.CountIf(ws.Range("A4:A" & ws.Rows.Count), cell.Value) > 1 Then
                    ws.Rows(cell.Row).ClearContents
                End If
            End If
        Next cell
        Application.EnableEvents = True
    End If

    If Not Intersect(Target, ws.Range("A4:A" & ws.Rows.Count)) Is Nothing Then
        
    Application.EnableEvents = False

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("C4:C" & lastRow).Formula = "=IF(LEN(A4)=17,LEFT(A4,5),IF(A4="""",""-"",""ERROR""))"
    ws.Range("E4:E" & lastRow).Formula = "=IF(LEN(A4)=17,MID(A4,6,6),IF(A4="""",""-"",""ERROR""))"
    ws.Range("G4:G" & lastRow).Formula = "=IF(LEN(A4)=17,MID(A4,12,2),IF(A4="""",""-"",""ERROR""))"
    ws.Range("H4:H" & lastRow).Formula = "=IF(LEN(A4)=17,MID(A4,14,4),IF(A4="""",""-"",""ERROR""))"

    ws.Range("C4:H" & lastRow).FormatConditions.Delete

    ApplyConditionalFormat ws.Range("C4:H" & lastRow)
        
        
    ' BUSQUEDA EN LA HOJA BD PARA CLIENTES ----------------------------------------------------------------------------------------
    
    For i = 4 To lastRow
        Dim numBusqueda As Variant
        numBusqueda = ws.Cells(i, 3).Value
        contenidoA = ws.Cells(i, 1).Value
    
        If Len(Trim(contenidoA)) = 0 Then
            ws.Cells(i, 4).Value = "-"
            ws.Cells(i, 4).Interior.ColorIndex = xlNone
        Else
            Dim encontrado As Range
            Set encontrado = Sheets("BD").Columns(1).Find(What:=numBusqueda, LookAt:=xlWhole)
    
            If Not encontrado Is Nothing Then
                ws.Cells(i, 4).Value = Sheets("BD").Cells(encontrado.Row, 2).Value
                ws.Cells(i, 4).Interior.ColorIndex = xlNone
            Else
                ws.Cells(i, 4).Value = "No se encontró cliente"
                ws.Cells(i, 4).Interior.Color = RGB(255, 0, 0)
            End If
        End If
    Next i
        
        Application.EnableEvents = True
    End If
    
    
    ' BUSCAR COINCIDENCIAS CON LA HOJA DE PEDIDOS ----------------------------------------------------------------------------------------
 
    Dim wsPedidos As Worksheet
    Dim lastRowPalletizado As Long
    Dim lastRowPedidos As Long
    Dim numPedido As Variant
    
    Set wsPedidos = ThisWorkbook.Sheets("PEDIDOS")

    lastRowPalletizado = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    lastRowPedidos = wsPedidos.Cells(wsPedidos.Rows.Count, "A").End(xlUp).Row

    Application.EnableEvents = False

    ' Iterar a través de las filas en PALLETIZADOR
    For i = 4 To lastRowPalletizado
        numPedido = ws.Cells(i, 5).Value

        If ws.Cells(i, 5).Value = "-" Then
            ws.Cells(i, 6).Value = "-"
            ws.Cells(i, 9).Value = 0
        Else
            Set encontrado = wsPedidos.Columns(1).Find(What:=numPedido, LookAt:=xlWhole)

            If Not encontrado Is Nothing Then
                ws.Cells(i, 6).Value = wsPedidos.Cells(encontrado.Row, 11).Value
                ws.Cells(i, 9).Value = wsPedidos.Cells(encontrado.Row, 21).Value * ws.Cells(i, 8).Value
            Else
                ws.Cells(i, 6).Value = "No se encontró descripción"
                ws.Cells(i, 9).Value = 0
                
                ws.Cells(i, 6).Interior.Color = RGB(255, 0, 0)
            End If
        End If
    Next i

    Application.EnableEvents = True
    
    ' AGREGAR MENU DESPLEGABLE "EN DEPOSITO - DESPACHADO" ----------------------------------------------------------------------------------------
    
    Set rngA = Intersect(Target, ws.Columns("A"))

    If rngA Is Nothing Then Exit Sub

    Application.EnableEvents = False

    For Each cellA In rngA

        Set cellJ = ws.Cells(cellA.Row, "K")
        
        If cellA.Value = "" Then
            cellJ.ClearContents
            cellJ.Interior.ColorIndex = -4142
        Else
            If cellJ.Value = "" Then
                cellJ.Value = "En Deposito"
            End If

            If cellJ.Value = "Despachado" Then
                cellJ.Interior.Color = RGB(0, 255, 0) ' Verde
            Else
                cellJ.Interior.Color = RGB(255, 255, 0) ' Amarillo
            End If
        End If
    Next cellA
    
    Application.EnableEvents = True
      
End Sub

' COLOR PARA LOS ERRORES ----------------------------------------------------------------------------------------

Sub ApplyConditionalFormat(rng As Range)
    Dim cell As Range
    For Each cell In rng
        If cell.Text = "ERROR" Then
            cell.Interior.Color = RGB(255, 0, 0) ' Rojo
        Else
            cell.Interior.ColorIndex = xlNone
        End If
    Next cell
End Sub

Private Sub btnEliminarFilasVacias_Click()
    Dim hojaPalletizado As Worksheet
    Dim ultimaFila As Long
    Dim fila As Long
    Dim filasBorradas As Boolean

    ' Definir la hoja de trabajo
    Set hojaPalletizado = ThisWorkbook.Sheets("PALLETIZADO")
    
    ultimaFila = hojaPalletizado.Cells(hojaPalletizado.Rows.Count, "A").End(xlUp).Row

    ' Inicializar la variable que indica si se han borrado filas
    filasBorradas = False

    For fila = ultimaFila To 4 Step -1
        If hojaPalletizado.Cells(fila, 1).Value = "" Then
            ' La celda en la columna A está vacía, eliminar toda la fila
            hojaPalletizado.Rows(fila).Delete
            filasBorradas = True
        End If
    Next fila

    If Not filasBorradas Then
        MsgBox "No se encontraron filas vacías para borrar.", vbInformation
    End If
End Sub







