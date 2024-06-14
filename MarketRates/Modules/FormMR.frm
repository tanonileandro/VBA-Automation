VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} COTIZACIONES 
   Caption         =   "EMBALAJES SRL - Carga de Cotizaciones"
   ClientHeight    =   12090
   ClientLeft      =   -1605
   ClientTop       =   -5805
   ClientWidth     =   10770
   OleObjectBlob   =   "FormMR.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "COTIZACIONES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnEnviarCorreo_Click()
    Dim OutlookApp As Object
    Dim OutlookMail As Object

    ' Crear instancia de Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    ' Configurar el correo electrónico
    With OutlookMail
        .To = "sistemas@embalajessrl.com.ar"
        .Subject = "Soporte Planilla Cotizaciones"
        .Body = ""
        ' .Attachments.Add "Ruta\archivo.pdf" ' Si quieres adjuntar archivos
        .Display ' Mostrar el correo para revisión
        ' .Send ' Utiliza .Send en lugar de .Display para enviar automáticamente
    End With

    ' Liberar recursos
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

End Sub



Private Sub CommandButton1_Click()
    Dim Numero As String
    Dim Nombre As String
    Dim mensaje As String
    Dim Hipervinculo As String
    
    ' Obtiene el número y el nombre del formulario de Excel
    ' Numero = InputBox("Ingrese el número de teléfono:")
    ' Nombre = InputBox("Ingrese el nombre:")
    Numero = "+5493417409678"
    Nombre = "Leandro"
    
    ' Construye el mensaje
    mensaje = "Buenos días " & Nombre & ", te envio este mensaje automático para solicitarte ayuda con el formulario de COTIZACIONES. Muchas gracias."
    
    ' Construye el hipervínculo
    Hipervinculo = "https://api.whatsapp.com/send?phone=" & Numero & "&text=" & mensaje
    
    ' Abre el hipervínculo
    ThisWorkbook.FollowHyperlink Hipervinculo
    
End Sub

Private Sub btnAyuda_Click()
MsgBox "Para obtener soporte, contactar de las siguientes maneras:" & vbCrLf & _
        " " & vbCrLf & _
        "- WhatsApp: +54 9 (341)-7409678" & vbCrLf & _
        "- Correo electrónico: sistemas@embalajessrl.com.ar" & vbCrLf & _
        " " & vbCrLf & _
        "Para enviar un WhatsApp o Correo rápidamente, haz clic en los botones ubicados a la derecha del ícono de ayuda." & vbCrLf & _
        " " & vbCrLf & _
        "(*) En caso de requerir reformas o modificaciones incuir todos los detalles.", vbInformation, "Ayuda"
End Sub

Private Sub ComboBoxClientes_Change()
    Dim nombreBuscado As String
    Dim i As Integer
    
    nombreBuscado = Trim(Me.ComboBoxClientes.Value)

    Me.ListBoxCoincidencias.Clear
    
    For i = 0 To Me.ComboBoxClientes.ListCount - 1
        If InStr(1, UCase(Me.ComboBoxClientes.List(i)), UCase(nombreBuscado), vbTextCompare) > 0 Then
            Me.ListBoxCoincidencias.AddItem Me.ComboBoxClientes.List(i)
        End If
    Next i
End Sub


Private Sub BtnBorrar_Click()

Load BorrarProduccion
BorrarProduccion.Show

End Sub

Private Sub BtnCancelar_Click()

Unload Me

End Sub

Private Sub BtnCargar_Click()
Application.ScreenUpdating = False
If Me.TxtIdCliente.Value <> Empty Then
    Dim Cell As Range
    Dim i As Single
    Dim j As Single
    Dim h As Single
    Dim NR As Long
    
    Set Cell = Worksheets("pedidos").Range("Tabla25")
    maximo = 1
    
    existe = 0
    For i = 2 To Cell.ListObject.Range.Rows.Count
    
        If Cell.ListObject.Range.Cells(i, 2) = Empty Then
            maximo = 1
        Else
        If Cell.ListObject.Range.Cells(i, 2) > maximo Then
            maximo = Cell.ListObject.Range.Cells(i, 2)
        End If
        End If

        
        If CStr(Me.TxtIdCliente.Value) = Cell.ListObject.Range.Cells(i, 3) And Cell.ListObject.Range.Cells(i, 1) = Date Then
            existe = 1
            num_cot = Cell.ListObject.Range.Cells(i, 2)
        End If
    Next i
    
          'CALCULO DE TARIFA X M2'
      Set Cell = Worksheets("TARIFARIO M2 2").Range("Tabla157912141620")
        For j = 1 To Cell.ListObject.Range.Rows.Count
            If Cell.ListObject.Range.Cells(j, 3) = Me.TxtCalidad.Value Then
                If Me.TxtCategoria.Value = "A" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(j, 4)
                    Else
                    If Me.TxtCategoria.Value = "B" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(j, 5)
                    Else
                    If Me.TxtCategoria.Value = "C" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(j, 6)
                    Else
                    If Me.TxtCategoria.Value = "D" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(j, 7)
                    Else
                    If Me.TxtCategoria.Value = "E" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(j, 8)
                    Else
                    If Me.TxtCategoria.Value = "F" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(j, 9)
                    End If
                    End If
                    End If
                    End If
                    End If
                End If
            End If
        Next j
            
      'CALCULO DE TARIFA DE FLETE X DESTINO'
      If OptionButtonSi.Value = True Then
        Set Cell = Worksheets("TARIFARIO FLETE 2").Range("TablaFlete")
        For h = 1 To Cell.ListObject.Range.Rows.Count
            If Cell.ListObject.Range.Cells(h, 3) = Me.ComboBoxDestino.Value Then
                tarifaflete = Cell.ListObject.Range.Cells(h, 6)
            End If
        Next h
      End If
      
        'CALCULO DE TARIFA DE FLETE X DESTINO'
           
         If existe = 1 Then
            Result2 = MsgBox("El cliente ya tiene una cotizacion cargada el día de hoy, desea AGREGAR otro registro de cotizacion?", vbOKCancel + vbQuestion, "Agregar registro")
             If Result2 = vbOK Then
                    If Me.TxtDescProd.Value = Empty And Me.ComboBoxDestino.Value = Empty And IsNumeric(Me.TxtCliente.Value) = True Then
                        MsgBox "Complete el articulo"
                        Else
                            If OptionButtonSi.Value = True Then
                        
                                Sheets("pedidos").Select
                                
                                NR = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
                                Cells(NR, 1) = Date
                                Cells(NR, 2) = num_cot
                                Cells(NR, 3) = Me.TxtIdCliente.Value
                                Cells(NR, 4) = Me.TxtDescCliente.Value
                                Cells(NR, 5) = Me.TxtCodProd.Value
                                Cells(NR, 6) = Me.TxtDescProd.Value
                                Cells(NR, 7) = "=[@[TARIFARIO $/M2]]*[@M2]"
                                Cells(NR, 9) = Me.TxtCalidad.Value
                                Cells(NR, 11) = CStr(Me.TxtM2.Value)
                                Cells(NR, 12) = "SI"
                                Cells(NR, 13) = Me.ComboBoxDestino.Value
                                Cells(NR, 14) = Me.TxtCategoria.Value
                                Cells(NR, 15) = tarifXm2 + tarifaflete
                            'carga cliente nuevo
                                Cells(NR, 16) = Me.TxtCategoria.Value
                                Cells(NR, 17) = Me.TxtDescProd.Value
                                Cells(NR, 18) = tarifXm2 + tarifaflete
                                Cells(NR, 20) = Me.TxtCondPago.Value
                                Cells(NR, 21) = Me.EjecutorCotizacion.Value
                            
                                Application.ScreenUpdating = True
                                MsgBox "Cotizacion Cargada"
                                
                                Sheets("pedidos").Select
                                Else
                                
                                Sheets("pedidos").Select

                                NR = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
                                Cells(NR, 1) = Date
                                Cells(NR, 2) = num_cot
                                Cells(NR, 3) = Me.TxtIdCliente.Value
                                Cells(NR, 4) = Me.TxtDescCliente.Value
                                Cells(NR, 5) = Me.TxtCodProd.Value
                                Cells(NR, 6) = Me.TxtDescProd.Value
                                Cells(NR, 7) = "=[@[TARIFARIO $/M2]]*[@M2]"
                                Cells(NR, 9) = Me.TxtCalidad.Value
                                Cells(NR, 11) = CStr(Me.TxtM2.Value)
                                Cells(NR, 12) = "NO"
                                Cells(NR, 14) = Me.TxtCategoria.Value
                                Cells(NR, 15) = tarifXm2 + tarifaflete
                            'carga cliente nuevo
                                Cells(NR, 16) = Me.TxtCategoria.Value
                                Cells(NR, 17) = Me.TxtDescProd.Value
                                Cells(NR, 18) = tarifXm2 + tarifaflete
                                Cells(NR, 20) = Me.TxtCondPago.Value
                                Cells(NR, 21) = Me.EjecutorCotizacion.Value
                                
                                Application.ScreenUpdating = True
                                MsgBox "Cotizacion Cargada"
                                
                                Sheets("pedidos").Select
                            End If
                            
                    End If
                End If
            Else
                            Sheets("pedidos").Select
                            If OptionButtonSi.Value = True Then
                                Sheets("pedidos").Select
                                NR = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
                                Cells(NR, 1) = Date
                                Cells(NR, 2) = maximo + 1
                                Cells(NR, 3) = Me.TxtIdCliente.Value
                                Cells(NR, 4) = Me.TxtDescCliente.Value
                                Cells(NR, 5) = Me.TxtCodProd.Value
                                Cells(NR, 6) = Me.TxtDescProd.Value
                                Cells(NR, 7) = "=[@[TARIFARIO $/M2]]*[@M2]"
                                Cells(NR, 9) = Me.TxtCalidad.Value
                                Cells(NR, 11) = CStr(Me.TxtM2.Value)
                                Cells(NR, 12) = "SI"
                                Cells(NR, 13) = Me.ComboBoxDestino.Value
                                Cells(NR, 14) = Me.TxtCategoria.Value
                                Cells(NR, 15) = tarifXm2 + tarifaflete
                            'carga cliente nuevo
                                Cells(NR, 16) = Me.TxtCategoria.Value
                                Cells(NR, 17) = Me.TxtDescProd.Value
                                Cells(NR, 18) = tarifXm2 + tarifaflete
                                Cells(NR, 20) = Me.TxtCondPago.Value
                                Cells(NR, 21) = Me.EjecutorCotizacion.Value
                                
                                Application.ScreenUpdating = True
                                MsgBox "Cotizacion Cargada"
                                
                                Sheets("pedidos").Select
                                Else
                                
                                Sheets("pedidos").Select
                                
                                NR = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
                                Cells(NR, 1) = Date
                                Cells(NR, 2) = maximo + 1
                                Cells(NR, 3) = Me.TxtIdCliente.Value
                                Cells(NR, 4) = Me.TxtDescCliente.Value
                                Cells(NR, 5) = Me.TxtCodProd.Value
                                Cells(NR, 6) = Me.TxtDescProd.Value
                                Cells(NR, 7) = "=[@[TARIFARIO $/M2]]*[@M2]"
                                Cells(NR, 9) = Me.TxtCalidad.Value
                                Cells(NR, 11) = CStr(Me.TxtM2.Value)
                                Cells(NR, 12) = "NO"
                                Cells(NR, 14) = Me.TxtCategoria.Value
                                Cells(NR, 15) = tarifXm2 + tarifaflete
                            'carga cliente nuevo
                                Cells(NR, 16) = Me.TxtCategoria.Value
                                Cells(NR, 17) = Me.TxtDescProd.Value
                                Cells(NR, 18) = tarifXm2 + tarifaflete
                                Cells(NR, 20) = Me.TxtCondPago.Value
                                Cells(NR, 21) = Me.EjecutorCotizacion.Value
                                
                                Application.ScreenUpdating = True
                                MsgBox "Cotizacion Cargada"
                                
                                Sheets("pedidos").Select
                            End If
                            
                            Sheets("pedidos").Select
        End If

        Else
        Application.ScreenUpdating = True
        MsgBox "Debe validar el cliente"
End If
End Sub
Private Sub BtnCotMasiva_Click()
    Application.ScreenUpdating = False
    If Me.TxtIdCliente.Value <> Empty Then
        Dim Cell As Range
        Dim Cell2 As Range
        Dim i As Single
        Dim j As Single
        Dim h As Single
        Dim g As Single
        Dim NR As Long
        Dim activeSheetName As String
        
        Set Cell = Worksheets("pedidos").Range("Tabla25")
        Set Cell2 = Worksheets("maestro articulos").Range("Tabla22")
        maximo = 1
        existe = 0
        
        For i = 2 To Cell.ListObject.Range.Rows.Count
            If Cell.ListObject.Range.Cells(i, 2) = Empty Then
                maximo = 1
            Else
                If Cell.ListObject.Range.Cells(i, 2) > maximo Then
                    maximo = Cell.ListObject.Range.Cells(i, 2)
                End If
            End If
        Next i
        
        ' Guardar el nombre de la hoja activa
        activeSheetName = ActiveSheet.Name
    
    Sheets("maestro articulos").Select
    For j = 1 To Cell2.ListObject.Range.Rows.Count
        If Val(Me.TxtIdCliente.Value) = Cell2.ListObject.Range.Cells(j, 1) Then
            cod_prod = Cell2.ListObject.Range.Cells(j, 2)
            desc_prod = Cell2.ListObject.Range.Cells(j, 12)
            calidad = Cell2.ListObject.Range.Cells(j, 13)
            m2 = Cell2.ListObject.Range.Cells(j, 21)
            
                   'CALCULO DE TARIFA X M2'
      Set Cell = Worksheets("TARIFARIO M2 2").Range("Tabla157912141620")
        For g = 1 To Cell.ListObject.Range.Rows.Count
            If Trim(Cell.ListObject.Range.Cells(g, 3)) = Trim(calidad) Then
                If Me.TxtCategoria.Value = "A" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(g, 4)
                    Else
                    If Me.TxtCategoria.Value = "B" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(g, 5)
                    Else
                    If Me.TxtCategoria.Value = "C" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(g, 6)
                    Else
                    If Me.TxtCategoria.Value = "D" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(g, 7)
                    Else
                    If Me.TxtCategoria.Value = "E" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(g, 8)
                    Else
                    If Me.TxtCategoria.Value = "F" Then
                    tarifXm2 = Cell.ListObject.Range.Cells(g, 9)
                    End If
                    End If
                    End If
                    End If
                    End If
                End If
            End If
        Next g

      'CALCULO DE TARIFA X M2'
            
      'CALCULO DE TARIFA DE FLETE X DESTINO'
      If OptionButtonSi.Value = True Then
        Set Cell = Worksheets("TARIFARIO FLETE 2").Range("TablaFlete")
        For h = 1 To Cell.ListObject.Range.Rows.Count
            If Cell.ListObject.Range.Cells(h, 3) = Me.ComboBoxDestino.Value Then
                tarifaflete = Cell.ListObject.Range.Cells(h, 6)
            End If
        Next h
      End If
        'CALCULO DE TARIFA DE FLETE X DESTINO'
               
                If Me.OptionButtonNo.Value = True Then
                                Sheets("pedidos").Select
                                NR = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
                                Cells(NR, 1) = Date
                                Cells(NR, 2) = maximo + 1
                                Cells(NR, 3) = Me.TxtIdCliente.Value
                                Cells(NR, 4) = Me.TxtDescCliente.Value
                                Cells(NR, 5) = cod_prod
                                Cells(NR, 6) = desc_prod
                                Cells(NR, 7) = "=[@[TARIFARIO $/M2]]*[@M2]"
                                Cells(NR, 9) = Trim(calidad)
                                Cells(NR, 11) = m2
                                Cells(NR, 12) = "NO"
                                Cells(NR, 14) = Me.TxtCategoria.Value
                                Cells(NR, 15) = tarifXm2 + tarifaflete
                                Cells(NR, 21) = Me.EjecutorCotizacion.Value
                                
                                Else
                                Sheets("pedidos").Select
                                NR = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
                                Cells(NR, 1) = Date
                                Cells(NR, 2) = maximo + 1
                                Cells(NR, 3) = Me.TxtIdCliente.Value
                                Cells(NR, 4) = Me.TxtDescCliente.Value
                                Cells(NR, 5) = cod_prod
                                Cells(NR, 6) = desc_prod
                                Cells(NR, 7) = "=[@[TARIFARIO $/M2]]*[@M2]"
                                Cells(NR, 9) = Trim(calidad)
                                Cells(NR, 11) = m2
                                Cells(NR, 12) = "SI"
                                Cells(NR, 13) = Me.ComboBoxDestino.Value
                                Cells(NR, 14) = Me.TxtCategoria.Value
                                Cells(NR, 15) = tarifXm2 + tarifaflete
                                Cells(NR, 21) = Me.EjecutorCotizacion.Value
                End If
                                
        End If
        Sheets("maestro articulos").Select
    Next j
    
    ' Volver a seleccionar la hoja activa
    Worksheets(activeSheetName).Select
    
    MsgBox "cotizacion masiva cargada"
End If
Application.ScreenUpdating = True
End Sub

Private Sub BtnValidar_Click()

    If Me.EjecutorCotizacion.Value = "" Then
        MsgBox "Antes de validar el cliente, por favor ingrese el nombre del Ejecutor de Ventas", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    TxtCategoria.Locked = False
    TxtCodProd.Locked = True
   
    
    Dim Cell As Range
    Dim i As Single
    Dim j As Single
    Dim k As Single
    Dim h As Single
    
    
    Set Cell = Worksheets("clientes activos").Range("Tabla10")
    
    existe = 0
    existe2 = 0
    
    For i = 1 To Cell.ListObject.Range.Rows.Count
    
        If CStr(Me.TxtClienteValidacion.Value) = Cell.ListObject.Range.Cells(i, 1) Then
           
            cuenta = Cell.ListObject.Range.Cells(i, 1)
            cliente = Cell.ListObject.Range.Cells(i, 2)
            pago = Cell.ListObject.Range.Cells(i, 14)
            categoria = Cell.ListObject.Range.Cells(i, 21)
            
            existe = 1
            MsgBox "CLIENTE ACTIVO ENCONTRADO"
        End If
            
    Next i
    
        Set Cell = Worksheets("pedidos").Range("Tabla25")
        For ñ = 1 To Cell.ListObject.Range.Rows.Count
            If CStr(Me.TxtClienteValidacion.Value) = Cell.ListObject.Range.Cells(ñ, 3) And Cell.ListObject.Range.Cells(ñ, 1) = Date Then
                
                cuenta = Cell.ListObject.Range.Cells(ñ, 3)
                cliente = Cell.ListObject.Range.Cells(ñ, 4)
                cotiz = Cell.ListObject.Range.Cells(ñ, 2)
                
                
                existe2 = 1
            End If
        Next ñ
        If existe2 = 1 Then
            MsgBox "CLIENTE CON COTIZACION ENCONTRADO"
        End If
            
        If existe = 1 Or existe2 = 1 Then
             
            Me.TxtIdCliente.Value = cuenta
            TxtIdCliente.Locked = True
            Me.TxtDescCliente.Value = cliente
            TxtDescCliente.Locked = True
            Me.TxtCondPago.Value = pago
            Txtcotiz.Locked = True
            Me.Txtcotiz.Value = cotiz
            Me.TxtCategoria.Value = categoria
            
            
            Set Cell = Worksheets("maestro articulos").Range("Tabla22")
            Me.ComboBoxProductos.Clear
    
            For j = 1 To Cell.ListObject.Range.Rows.Count
    
                If CStr(cuenta) = CStr(Cell.ListObject.Range.Cells(j, 1)) Then ' If cuenta = Cell.ListObject.Range.Cells(j, 1) Then  se modifico porque no coincidian los tipos, es necesario que sean string
                    Me.ComboBoxProductos.AddItem (Trim(Cell.ListObject.Range.Cells(j, 12)) + "  -  " + "(" + CStr(Trim(Cell.ListObject.Range.Cells(j, 2)) + ")"))
                End If
        
            Next j

        Else
        Application.ScreenUpdating = True
        result3 = MsgBox("Cliente no encontrado,desea generar una cotizacion a un nuevo cliente?", vbOKCancel + vbQuestion, "Agregar registro")
      
    
        If result3 = vbOK Then
    Set Cell = Worksheets("pedidos").Range("Tabla25")
    cuenta_CN = 0
        
        For h = 1 To Cell.ListObject.Range.Rows.Count
            If IsNumeric(Cell.ListObject.Range.Cells(h, 3)) = False And Cell.ListObject.Range.Cells(h, 1) = Date And existe_CN = 0 Then
            cuenta_CN = cuenta_CN + 1
            End If
        Next h

            Me.TxtIdCliente.Value = "CLIENTE NUEVO" + " " + CStr(cuenta_CN)
            TxtIdCliente.Locked = True
            Me.TxtCodProd.Value = "-"
            TxtCodProd.Locked = True
            Me.TxtCliente.Value = "1"
            TxtDescCliente.Locked = False
            TxtCondPago.Locked = False
            TxtCondPago.Locked = False
            TxtM2.Locked = False
            TxtCalidad.Locked = False
            TxtCategoria.Locked = False
        End If
    End If
    
        ' Encontrar la última fila con datos en la columna B de la hoja "PEDIDOS"
    Dim lastRow As Long
    With Worksheets("PEDIDOS")
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
    End With
    
    ' Obtener el valor de la celda en la última fila de la columna B y sumarle uno
    Dim nextNumber As Long
    If lastRow >= 2 Then
        nextNumber = CLng(Worksheets("PEDIDOS").Cells(lastRow, "B").Value) + 1
    Else
        nextNumber = 1 ' Si no hay datos, comenzar desde 1
    End If
    
    ' Mostrar el próximo número en el campo Txtcotiz
    Me.Txtcotiz.Value = nextNumber
    Me.Txtcotiz.Locked = True
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub LblDescripcionTexto_Click()

End Sub

Private Sub Lblordentexto_Click()

End Sub


Private Sub ComboBoxProductos_Change()
Dim Cell As Range
Dim i As Single


Set Cell = Worksheets("maestro articulos").Range("Tabla22")

For i = 1 To Cell.ListObject.Range.Rows.Count

    If Me.ComboBoxProductos.Text = Trim(Cell.ListObject.Range.Cells(i, 12)) + "  -  " + "(" + CStr(Trim(Cell.ListObject.Range.Cells(i, 2)) + ")") Then
        
        codigo = Cell.ListObject.Range.Cells(i, 2)
        producto = Cell.ListObject.Range.Cells(i, 12)
        calidad = Cell.ListObject.Range.Cells(i, 13)
        m2 = Cell.ListObject.Range.Cells(i, 21)
        Me.TxtCodProd.Value = codigo
        TxtCodProd.Locked = True
        Me.TxtCalidad.Value = Trim(calidad)
        Me.TxtM2.Value = Str(m2)
        TxtM2.Locked = True
        Me.TxtDescProd.Value = producto
        TxtDescProd.Locked = True
     
    End If
    
Next i
TxtCategoria.Locked = False
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub ListCoincidencias_Change()

End Sub



Private Sub Label22_Click()

End Sub

Private Sub ListBoxCoincidencias_Click()
    If Me.ListBoxCoincidencias.ListIndex >= 0 Then
        Dim selectedItem As String
        selectedItem = Me.ListBoxCoincidencias.List(Me.ListBoxCoincidencias.ListIndex)
        
        Dim parts() As String
        parts = Split(selectedItem, "-")
        Me.TxtCliente.Value = Trim(parts(1))
        Me.TxtClienteValidacion.Value = Trim(parts(0))
    End If
End Sub

Private Sub NombreSolicitante_Change()
    Dim nombreSeleccionada As String
    nombreSeleccionada = Trim(Me.NombreSolicitante.Value)
End Sub

Private Sub OptionButtonNo_Click()
If OptionButtonNo.Value = True Then
    Me.ComboBoxDestino.Clear
End If
End Sub

Private Sub OptionButtonSi_Click()
If OptionButtonSi.Value = True And Me.TxtIdCliente.Value <> Empty Then
 Dim j As Single
        Set Cell = Worksheets("TARIFARIO FLETE 2").Range("TablaFlete")
        Me.ComboBoxDestino.Clear

        For j = 1 To Cell.ListObject.Range.Rows.Count
    
            Me.ComboBoxDestino.AddItem (Cell.ListObject.Range.Cells(j, 3))
    
        Next j
End If
End Sub

Private Sub TxtBusqueda_Change()
    Dim nombreBuscado As String
    Dim i As Integer
    Dim coincidencias As String
    
    nombreBuscado = UCase(Trim(Me.TxtBusqueda.Value))
    coincidencias = ""
    
    Me.ListCoincidencias.Clear
    
    For i = 0 To Me.ComboBoxClientes.ListCount - 1
        If InStr(1, UCase(Me.ComboBoxClientes.List(i)), nombreBuscado, vbTextCompare) > 0 Then
            coincidencias = coincidencias & Me.ComboBoxClientes.List(i) & vbCrLf
            Me.ListCoincidencias.AddItem Me.ComboBoxClientes.List(i)
        End If
    Next i
End Sub

Private Sub EjecutorCotizacion_Change()

End Sub

Private Sub TxtCalidad_Change()

End Sub

Private Sub TxtCliente_Change()

End Sub

Private Sub Txtcotiz_Change()

End Sub

Private Sub UserForm_Initialize()
    
    Dim lastRow As Long
    With Worksheets("PEDIDOS")
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
    End With
    
    Dim nextNumber As Long
    If lastRow >= 2 Then
        nextNumber = CLng(Worksheets("PEDIDOS").Cells(lastRow, "B").Value) + 1
    Else
        nextNumber = 1 ' Si no hay datos, comenzar desde 1
    End If
    
    Me.Txtcotiz.Value = nextNumber
    
    Dim j As Long
    Dim Cell As Range
    
    With Worksheets("clientes activos")
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Set Cell = .Range("A2:B" & lastRow)
    End With
    
    ' Limpiar el ComboBoxClientes
    Me.ComboBoxClientes.Clear
    
    ' Agregar elementos al ComboBoxClientes
    For j = 1 To Cell.Rows.Count
        Me.ComboBoxClientes.AddItem Trim(Cell.Cells(j, 1).Value) & " - " & Trim(Cell.Cells(j, 2).Value)
    Next j
    
    Me.ListBoxCoincidencias.Clear
    
    ' Agregar elementos al ListBoxCoincidencias
    For j = 1 To Cell.Rows.Count
        Me.ListBoxCoincidencias.AddItem Trim(Cell.Cells(j, 1).Value) & " - " & Trim(Cell.Cells(j, 2).Value)
    Next j
    
    Me.Width = 540
    Me.Height = 502
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
End Sub


