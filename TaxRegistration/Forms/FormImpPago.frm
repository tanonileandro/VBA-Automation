VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormImpPago 
   Caption         =   "EMBALAJES SRL"
   ClientHeight    =   6720
   ClientLeft      =   750
   ClientTop       =   2040
   ClientWidth     =   8775.001
   OleObjectBlob   =   "FormImpPago.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FormImpPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pdfImpPath As String
Private pdfPagoPath As String

Private Sub UserForm_Initialize()
    InitializeForm
End Sub

Private Sub InitializeForm()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim uniqueValues As Collection
    Dim item As Variant
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    
    ' Configurar la hoja y la tabla
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("ImpAnual")
    Set tbl = ws.ListObjects("Tabla3")

    ' Verifica que la tabla y la hoja son correctas
    If ws Is Nothing Then
        MsgBox "La hoja 'ImpAnual' no se encontró.", vbCritical
        Exit Sub
    End If
    If tbl Is Nothing Then
        MsgBox "La tabla 'Tabla3' no se encontró en la hoja 'ImpAnual'.", vbCritical
        Exit Sub
    End If
    
    Me.ShowLinkImp.Enabled = False
    Me.ShowMonto.Enabled = False
    Me.ShowFecha.Enabled = False
    Me.ShowLinkPago.Enabled = False

    ' Crear una colección para almacenar valores únicos
    Set uniqueValues = New Collection
    
    startRow = tbl.Range.row + 1
    endRow = tbl.Range.Rows.Count + tbl.Range.row - 1
    
    On Error Resume Next
    For i = startRow To endRow
        Set cell = ws.Cells(i, 4)
        If cell.Value <> "" Then
            uniqueValues.Add cell.Value, CStr(cell.Value)
        End If
    Next i
    On Error GoTo 0

    If uniqueValues.Count = 0 Then
        MsgBox "No se encontraron valores únicos en la columna C de la tabla 'Tabla3'.", vbInformation
        Exit Sub
    End If

    For Each item In uniqueValues
        Me.ServiceType.AddItem item
    Next item
    
    ' Tamaño del formulario
    Me.Width = 453
    Me.Height = 300
    
    ' Posición del formulario al centro de la pantalla principal de Excel
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    
Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
End Sub

Private Sub FilterServiceType(ByVal month As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim uniqueValues As Collection
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim item As Variant

    Me.ServiceType.Clear

    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("ImpAnual")
    Set tbl = ws.ListObjects("Tabla3")

    If ws Is Nothing Then
        MsgBox "La hoja 'ImpAnual' no se encontró.", vbCritical
        Exit Sub
    End If
    If tbl Is Nothing Then
        MsgBox "La tabla 'Tabla3' no se encontró en la hoja 'ImpAnual'.", vbCritical
        Exit Sub
    End If

    Set uniqueValues = New Collection
    
    startRow = tbl.HeaderRowRange.row + 1
    endRow = tbl.Range.Rows.Count + tbl.HeaderRowRange.row - 1
    
    On Error Resume Next
    For i = startRow To endRow
        Set cell = ws.Cells(i, 1) ' Columna A
        If cell.Value = month Then
            uniqueValues.Add cell.Offset(0, 3).Value, CStr(cell.Offset(0, 3).Value) ' Columna C
        End If
    Next i
    On Error GoTo 0

    If uniqueValues.Count = 0 Then
        MsgBox "No se encontraron SERVICIOS o IMPUESTOS para el mes seleccionado.", vbInformation
        Exit Sub
    End If

    For Each item In uniqueValues
        Me.ServiceType.AddItem item
    Next item

Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
End Sub

' BOTON CARGAR FORMULARIO

Private Sub BtnCargar_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim SelectedDetail As String
    Dim startRow As Long
    Dim endRow As Long
    Dim i As Long
    Dim missingInfo As String
    Dim proceed As VbMsgBoxResult

    ' Obtener el valor seleccionado en ServiceDetail
    SelectedDetail = Me.ServiceDetail.Value
    
    ' Verificar que se haya seleccionado un detalle de servicio
    If SelectedDetail = "" Then
        MsgBox "Por favor, selecciona un detalle de servicio.", vbInformation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("ImpAnual")
    Set tbl = ws.ListObjects("Tabla3")

    If ws Is Nothing Then
        MsgBox "La hoja 'ImpAnual' no se encontró.", vbCritical
        Exit Sub
    End If
    
    If tbl Is Nothing Then
        MsgBox "La tabla 'Tabla3' no se encontró en la hoja 'ImpAnual'.", vbCritical
        Exit Sub
    End If

   
    startRow = tbl.HeaderRowRange.row + 1
    endRow = tbl.Range.Rows.Count + tbl.HeaderRowRange.row - 1
    
    On Error Resume Next
    For i = startRow To endRow
        Set cell = ws.Cells(i, 5)
        If cell.Value = SelectedDetail Then
            missingInfo = ""
            
            ' Verificar campos y construir el mensaje
            If pdfImpPath = "" Then
                missingInfo = missingInfo & "Impuesto, "
            End If
            If pdfPagoPath = "" Then
                missingInfo = missingInfo & "Link de Pago, "
            End If
            If Me.TextBoxMonto.Value = "" Then
                missingInfo = missingInfo & "Monto, "
            End If
            If Me.TextBoxFechaPago.Value = "" Then
                missingInfo = missingInfo & "Fecha de Pago, "
            End If
            
            If Len(missingInfo) > 0 Then
                missingInfo = Left(missingInfo, Len(missingInfo) - 2)
            End If
            
            If missingInfo <> "" Then
                proceed = MsgBox("Faltan los siguientes campos: " & missingInfo & ". ¿Desea continuar?", vbYesNo + vbExclamation, "Campos incompletos")
                If proceed = vbNo Then
                    Exit Sub
                End If
            End If

            If pdfImpPath <> "" Then
                InsertPDF ws, cell.Offset(0, 7), "Abrir Comprobante", pdfImpPath
            End If
            
            If pdfPagoPath <> "" Then
                InsertPDF ws, cell.Offset(0, 10), "Abrir Comprobante", pdfPagoPath
            End If
            
            If Me.TextBoxMonto.Value <> "" Then
                cell.Offset(0, 8).Value = Me.TextBoxMonto.Value ' Columna M
            End If
            
            If Me.TextBoxFechaPago.Value <> "" Then
                cell.Offset(0, 9).Value = Me.TextBoxFechaPago.Value ' Columna N
            End If
            
            Exit For
        End If
    Next i
    On Error GoTo 0

    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
End Sub

Private Sub ButtonCargaImp_Click()
    ' Abrir el diálogo para seleccionar el PDF de Imp
    pdfImpPath = SelectPDFFile
End Sub

Private Sub ButtonCargaPago_Click()
    ' Abrir el diálogo para seleccionar el PDF de Pago
    pdfPagoPath = SelectPDFFile
End Sub

Private Function SelectPDFFile() As String
    Dim fd As FileDialog
    Dim selectedFile As String

    ' Configurar el diálogo de selección de archivo
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleccionar archivo PDF"
        .Filters.Clear
        .Filters.Add "Archivos PDF", "*.pdf"
        .FilterIndex = 1
        .ButtonName = "Seleccionar"
        .AllowMultiSelect = False

        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
        End If
    End With

    Set fd = Nothing
    SelectPDFFile = selectedFile
End Function

Private Sub InsertPDF(ws As Worksheet, targetCell As Range, displayText As String, pdfPath As String)
    ' Insertar el PDF en la celda especificada con el texto de visualización
    If pdfPath <> "" Then
        ws.Hyperlinks.Add Anchor:=targetCell, Address:=pdfPath, TextToDisplay:=displayText
    End If
End Sub
Private Sub CloseForm_Click()
    FormImpPago.Hide
End Sub

Private Sub OptionAbr_Click()
    FilterServiceType "abr"
End Sub

Private Sub OptionAgo_Click()
    FilterServiceType "ago"
End Sub

Private Sub OptionDic_Click()
    FilterServiceType "dic"
End Sub

Private Sub OptionEne_Click()
    FilterServiceType "ene"
End Sub

Private Sub OptionFeb_Click()
    FilterServiceType "feb"
End Sub

Private Sub OptionJul_Click()
    FilterServiceType "jul"
End Sub

Private Sub OptionJun_Click()
    FilterServiceType "jun"
End Sub

Private Sub OptionMar_Click()
    FilterServiceType "mar"
End Sub

Private Sub OptionMay_Click()
    FilterServiceType "may"
End Sub

Private Sub OptionNov_Click()
    FilterServiceType "nov"
End Sub

Private Sub OptionOct_Click()
    FilterServiceType "oct"
End Sub

Private Sub OptionSep_Click()
    FilterServiceType "sep"
End Sub

Private Sub SelectedType_Change()
    ' No se necesita hacer nada aquí en este momento
    Me.SelectedType.Enabled = False
End Sub

Private Sub ServiceDetail_Change()
    Dim selectedLinkImp As String
    Dim selectedMonto As String
    Dim selectedFecha As String
    Dim selectedLinkPago As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim SelectedDetail As String
    Dim startRow As Long
    Dim endRow As Long
    Dim i As Long

    ' Limpiar los TextBoxes
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxObservaciones.Value = ""

    ' Obtener el valor seleccionado en ServiceDetail
    SelectedDetail = Me.ServiceDetail.Value
    
    ' Verificar que se haya seleccionado un detalle de servicio
    If SelectedDetail = "" Then
        MsgBox "Por favor, selecciona un detalle de servicio.", vbInformation
        Exit Sub
    End If

    ' Configurar la hoja y la tabla
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("ImpAnual")
    Set tbl = ws.ListObjects("Tabla3")

    ' Verifica que la tabla y la hoja son correctas
    If ws Is Nothing Then
        MsgBox "La hoja 'ImpAnual' no se encontró.", vbCritical
        Exit Sub
    End If
    
    If tbl Is Nothing Then
        MsgBox "La tabla 'Tabla3' no se encontró en la hoja 'ImpAnual'.", vbCritical
        Exit Sub
    End If

    startRow = tbl.HeaderRowRange.row + 1
    endRow = tbl.Range.Rows.Count + tbl.HeaderRowRange.row - 1
    
    On Error Resume Next
    For i = startRow To endRow
        Set cell = ws.Cells(i, 5) ' Columna D
        If cell.Value = SelectedDetail Then
            Me.TextBoxCuenta.Value = cell.Offset(0, 2).Value ' Columna G
            Me.TextBoxCuenta.Enabled = False
            Me.TextBoxObservaciones.Value = cell.Offset(0, 11).Value ' Columna P
            
            ' Almacenar valores en variables públicas
            selectedLinkImp = cell.Offset(0, 7).Value ' Columna L
            selectedMonto = cell.Offset(0, 8).Value ' Columna M
            selectedFecha = cell.Offset(0, 9).Value ' Columna N
            selectedLinkPago = cell.Offset(0, 10).Value ' Columna O
            
            ' Mostrar los valores en los campos correspondientes
            Me.ShowLinkImp.Text = selectedLinkImp
            Me.ShowMonto.Text = selectedMonto
            Me.ShowFecha.Text = selectedFecha
            Me.ShowLinkPago.Text = selectedLinkPago
            
            ' Configurar los campos como de solo lectura
            Me.ShowLinkImp.Enabled = False
            Me.ShowMonto.Enabled = False
            Me.ShowFecha.Enabled = False
            Me.ShowLinkPago.Enabled = False
            
            ' Llamar al evento TextBoxInfo_Change para verificar campos automáticamente
            TextBoxInfo_Change
            Exit For
        End If
    Next i
    On Error GoTo 0

    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
End Sub

Private Sub TextBoxInfo_Change()
    Dim missingInfo As String
    
    ' Verificar si alguno de los campos está vacío
    If Me.ShowLinkImp.Text = "" Or Me.ShowMonto.Text = "" Or Me.ShowFecha.Text = "" Or Me.ShowLinkPago.Text = "" Then
        missingInfo = "Falta cargar: "
        If Me.ShowLinkImp.Text = "" Then missingInfo = missingInfo & "Link del impuesto"
        If Me.ShowMonto.Text = "" Then
            If Len(missingInfo) > Len("Falta cargar: ") Then
                missingInfo = missingInfo & ", Monto"
            Else
                missingInfo = missingInfo & "Monto"
            End If
        End If
        If Me.ShowFecha.Text = "" Then
            If Len(missingInfo) > Len("Falta cargar: ") Then
                missingInfo = missingInfo & ", Fecha de pago"
            Else
                missingInfo = missingInfo & "fecha de pago"
            End If
        End If
        If Me.ShowLinkPago.Text = "" Then
            If Len(missingInfo) > Len("Falta cargar: ") Then
                missingInfo = missingInfo & ", Link de Pago"
            Else
                missingInfo = missingInfo & "link de Pago"
            End If
        End If
        
        Me.TextBoxInfo.Text = missingInfo
    Else
        ' Mostrar mensaje de todos los campos completos
        Me.TextBoxInfo.Text = "Pago cargado anteriormente, todos sus campos estan completos"
    End If
End Sub
Private Sub ServiceType_Change()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim SelectedType As String
    Dim startRow As Long
    Dim endRow As Long
    Dim i As Long

    ' Limpiar el ComboBox ServiceDetail
    Me.ServiceDetail.Clear

    ' Obtener el valor seleccionado en ServiceType
    SelectedType = Me.ServiceType.Value
    
    ' Verificar que se haya seleccionado un tipo de servicio
    If SelectedType = "" Then
        MsgBox "Por favor, selecciona un tipo de servicio.", vbInformation
        Exit Sub
    End If

    ' Mostrar el valor seleccionado en SelectedType
    Me.SelectedType.Value = SelectedType

    ' Configurar la hoja y la tabla
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("ImpAnual")
    Set tbl = ws.ListObjects("Tabla3")

    ' Verifica que la tabla y la hoja son correctas
    If ws Is Nothing Then
        MsgBox "La hoja 'ImpAnual' no se encontró.", vbCritical
        Exit Sub
    End If
    
    If tbl Is Nothing Then
        MsgBox "La tabla 'Tabla3' no se encontró en la hoja 'ImpAnual'.", vbCritical
        Exit Sub
    End If

    startRow = tbl.HeaderRowRange.row + 1
    endRow = tbl.Range.Rows.Count + tbl.HeaderRowRange.row - 1
    
    On Error Resume Next
    For i = startRow To endRow
        Set cell = ws.Cells(i, 4) ' Columna C
        If cell.Value = SelectedType Then
            Me.ServiceDetail.AddItem cell.Offset(0, 1).Value ' Columna D
        End If
    Next i
    On Error GoTo 0

    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
End Sub
