VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormImpPago 
   Caption         =   "EMBALAJES SRL"
   ClientHeight    =   14475
   ClientLeft      =   1740
   ClientTop       =   5550
   ClientWidth     =   16995
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
    
    Me.ShowLinkImp.Enabled = False
    Me.ShowLinkPago.Enabled = False
    Me.ShowFecha.Enabled = False
    Me.ShowFechaVto.Enabled = False
    Me.ShowMonto.Enabled = False
    Me.ShowObservaciones.Enabled = False
    Me.TextBoxInfo.Locked = True
    Me.TextBoxCuenta.Enabled = False
    Me.SelectedType.Enabled = False
    
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
    Me.Width = 450
    Me.Height = 385
    
    ' Posición del formulario al centro de la pantalla principal de Excel
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    
Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error: " & Err.Description, vbCritical
End Sub

Private Sub ServiceType_Change()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim SelectedType As String
    Dim SelectedMonth As String
    Dim startRow As Long
    Dim endRow As Long
    Dim i As Long

    ' Limpiar el ComboBox ServiceDetail
    Me.ServiceDetail.Clear

    ' Obtener el valor seleccionado en ServiceType
    SelectedType = Me.ServiceType.Value

    ' Obtener el valor seleccionado en SelectedType (mes)
    SelectedMonth = Me.SelectedType.Value

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

    startRow = tbl.HeaderRowRange.row + 1
    endRow = tbl.Range.Rows.Count + tbl.HeaderRowRange.row - 1

    On Error Resume Next
    For i = startRow To endRow
        Set cell = ws.Cells(i, 4) ' Columna D (Tipo de Servicio)
        If ws.Cells(i, 1).Value = SelectedMonth And cell.Value = SelectedType Then
            Me.ServiceDetail.AddItem ws.Cells(i, 5).Value ' Columna E (Detalle del Servicio)
        End If
    Next i
    On Error GoTo 0
End Sub
Private Sub ServiceDetail_Change()
    Dim selectedLinkImp As String
    Dim selectedMonto As String
    Dim selectedFecha As String
    Dim selectedFechaVto As String
    Dim selectedObservaciones As String
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

    startRow = tbl.HeaderRowRange.row + 1
    endRow = tbl.Range.Rows.Count + tbl.HeaderRowRange.row - 1
    
    On Error Resume Next
    For i = startRow To endRow
        Set cell = ws.Cells(i, 5) ' Columna D
        If cell.Value = SelectedDetail Then
            Me.TextBoxCuenta.Value = cell.Offset(0, 2).Value ' Columna G
            Me.TextBoxFechaVto.Value = cell.Offset(0, 6).Value ' Columna P
            Me.TextBoxFechaPago.Value = cell.Offset(0, 9).Value ' Columna P
            Me.TextBoxMonto.Value = cell.Offset(0, 8).Value ' Columna P
            Me.TextBoxObservaciones.Value = cell.Offset(0, 11).Value ' Columna P
            
            ' Almacenar valores en variables públicas
            selectedFechaVto = cell.Offset(0, 6).Value ' Columna M
            selectedLinkImp = cell.Offset(0, 7).Value ' Columna L
            selectedMonto = cell.Offset(0, 8).Value ' Columna M
            selectedFecha = cell.Offset(0, 9).Value ' Columna N
            selectedLinkPago = cell.Offset(0, 10).Value ' Columna O
            selectedObservaciones = cell.Offset(0, 11).Value ' Columna O
            
            ' Mostrar los valores en los campos correspondientes
            Me.ShowLinkImp.Text = selectedLinkImp
            Me.ShowMonto.Text = selectedMonto
            Me.ShowFecha.Text = selectedFecha
            Me.ShowFechaVto.Text = selectedFechaVto
            Me.ShowLinkPago.Text = selectedLinkPago
            Me.ShowObservaciones.Text = selectedObservaciones
            
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
Private Sub SelectedType_Change()
    ' No se necesita hacer nada aquí en este momento
    Me.SelectedType.Enabled = False
End Sub

Private Sub OptionEne_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "ene"
    ServiceType_Change
End Sub

Private Sub OptionFeb_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "feb"
    ServiceType_Change
End Sub
Private Sub OptionMar_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "mar"
    ServiceType_Change
End Sub
Private Sub OptionAbr_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""

    ' Actualizar y filtrar por el mes de abril
    Me.SelectedType.Value = "abr"
    ServiceType_Change
End Sub
Private Sub OptionMay_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "may"
    ServiceType_Change
End Sub
Private Sub OptionJun_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "jun"
    ServiceType_Change
End Sub
Private Sub OptionJul_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "jul"
    ServiceType_Change
End Sub
Private Sub OptionAgo_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "ago"
    ServiceType_Change
End Sub
Private Sub OptionSep_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "sep"
    ServiceType_Change
End Sub
Private Sub OptionOct_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "oct"
    ServiceType_Change
End Sub
Private Sub OptionNov_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "nov"
    ServiceType_Change
End Sub
Private Sub OptionDic_Click()
    ' Limpiar campos relevantes
    Me.SelectedType.Value = ""
    Me.ServiceDetail.Clear
    Me.TextBoxCuenta.Value = ""
    Me.TextBoxFechaVto.Value = ""
    Me.TextBoxFechaPago.Value = ""
    Me.TextBoxMonto.Value = ""
    Me.TextBoxObservaciones.Value = ""
    Me.ShowLinkImp.Text = ""
    Me.ShowMonto.Text = ""
    Me.ShowFecha.Text = ""
    Me.ShowFechaVto.Text = ""
    Me.ShowLinkPago.Text = ""
    Me.ShowObservaciones.Text = ""
    
    Me.SelectedType.Value = "dic"
    ServiceType_Change
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

    startRow = tbl.HeaderRowRange.row + 1
    endRow = tbl.Range.Rows.Count + tbl.HeaderRowRange.row - 1
    
    On Error Resume Next
    For i = startRow To endRow
        Set cell = ws.Cells(i, 5)
        If cell.Value = SelectedDetail Then
            missingInfo = ""
            
            ' Verificar campos y construir el mensaje
            If pdfImpPath = "" Then
                missingInfo = missingInfo & "Link Impuesto, "
            End If
            If pdfPagoPath = "" Then
                missingInfo = missingInfo & "Link Pago, "
            End If
            If Me.TextBoxFechaVto.Value = "" Then
                missingInfo = missingInfo & "Fecha Vto, "
            End If
            If Me.TextBoxFechaPago.Value = "" Then
                missingInfo = missingInfo & "Fecha de Pago, "
            End If
            If Me.TextBoxMonto.Value = "" Then
                missingInfo = missingInfo & "Monto, "
            End If
            If Me.TextBoxObservaciones.Value = "" Then
                missingInfo = missingInfo & "Observaciones, "
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
            
            If Me.TextBoxFechaVto.Text <> "" Then
                cell.Offset(0, 6).Value = "'" & Me.TextBoxFechaVto.Text ' Columna N
            End If
            If Me.TextBoxFechaPago.Text <> "" Then
                cell.Offset(0, 9).Value = "'" & Me.TextBoxFechaPago.Text ' Columna N
            End If
            If Me.TextBoxMonto.Value <> "" Then
                cell.Offset(0, 8).Value = Me.TextBoxMonto.Value
            End If
            If Me.TextBoxObservaciones.Value <> "" Then
                cell.Offset(0, 11).Value = Me.TextBoxObservaciones.Value
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
    
    ' Mostrar mensaje en el label MsjCargaImp si se ha seleccionado un PDF
    If pdfImpPath <> "" Then
        Me.MsjCargaImp.Caption = "PDF cargado OK"
    End If
End Sub

Private Sub ButtonCargaPago_Click()
    ' Abrir el diálogo para seleccionar el PDF de Pago
    pdfPagoPath = SelectPDFFile
    
    ' Mostrar mensaje en el label MsjCargaPago si se ha seleccionado un PDF
    If pdfPagoPath <> "" Then
        Me.MsjCargaPago.Caption = "PDF cargado OK"
    End If
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
        Else
            ' Devolver un valor nulo si el usuario cancela la selección
            SelectPDFFile = ""
            Exit Function
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

Private Sub TextBoxInfo_Change()
    Dim missingInfo As String
    
    ' Verificar si alguno de los campos está vacío
    If Me.ShowLinkImp.Text = "" Or Me.ShowMonto.Text = "" Or Me.ShowFecha.Text = "" Or Me.ShowFechaVto.Text = "" Or Me.ShowObservaciones.Text = "" Or Me.ShowLinkPago.Text = "" Then
        missingInfo = "Falta cargar: "
        If Me.ShowLinkImp.Text = "" Then missingInfo = missingInfo & "Link Impuesto"
        If Me.ShowFechaVto.Text = "" Then
        If Me.ShowLinkPago.Text = "" Then
        If Len(missingInfo) > Len("Falta cargar: ") Then
                missingInfo = missingInfo & ", Link Pago"
            Else
                missingInfo = missingInfo & "Link Pago"
            End If
        End If
        If Len(missingInfo) > Len("Falta cargar: ") Then
                missingInfo = missingInfo & ", Fecha Vto"
            Else
                missingInfo = missingInfo & "Fecha Vto"
            End If
        End If
        If Me.ShowFecha.Text = "" Then
            If Len(missingInfo) > Len("Falta cargar: ") Then
                missingInfo = missingInfo & ", Fecha Pago"
            Else
                missingInfo = missingInfo & "Fecha Pago"
            End If
        End If
        If Me.ShowMonto.Text = "" Then
            If Len(missingInfo) > Len("Falta cargar: ") Then
                missingInfo = missingInfo & ", Monto"
            Else
                missingInfo = missingInfo & "Monto"
            End If
        End If
        If Me.ShowObservaciones.Text = "" Then
            If Len(missingInfo) > Len("Falta cargar: ") Then
                missingInfo = missingInfo & ", Observaciones"
            Else
                missingInfo = missingInfo & "Observaciones"
            End If
        End If
        
        Me.TextBoxInfo.Text = missingInfo
    Else
        ' Mostrar mensaje de todos los campos completos
        Me.TextBoxInfo.Text = "Pago cargado anteriormente, todos sus campos estan completos"
    End If
End Sub
