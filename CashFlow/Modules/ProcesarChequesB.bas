Attribute VB_Name = "ProcesarChequesB"
Sub ProcesarChequesB()

    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False

    Dim rutaPlanilla As String
    rutaPlanilla = "Y:\PROVEEDORES\PAGO A PROVEEDORES\Planilla_Pagos_2024.xlsm"

    Dim planillaPagos As Workbook
    On Error Resume Next
    Set planillaPagos = Workbooks("Planilla_Pagos_2024.xlsm")
    On Error GoTo 0

    If planillaPagos Is Nothing Then
        Set planillaPagos = Workbooks.Open(rutaPlanilla)
        planillaPagos.Windows(1).Visible = False
    End If

    planillaPagos.Windows(1).Visible = False

    Dim hojaMensual As Worksheet
    Set hojaMensual = ThisWorkbook.Sheets("Mensual")

    Dim sumaValores As Double
    sumaValores = 0

    Dim hojaProveedores As Worksheet
    On Error Resume Next
    Set hojaProveedores = planillaPagos.Sheets("PROVEEDORES")
    On Error GoTo 0

    If Not hojaProveedores Is Nothing Then
        sumaValores = SumarValoresCheques(hojaProveedores)
    End If

    hojaMensual.Range("F23").Value = Abs(sumaValores)

    If Not planillaPagos.Windows(1).Visible Then
        planillaPagos.Windows(1).Visible = True
    End If

    planillaPagos.Close SaveChanges:=False
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
End Sub

Function SumarValoresCheques(hoja As Worksheet) As Double
    Dim suma As Double
    Dim fila As Long
    For fila = 2 To hoja.Cells(hoja.Rows.Count, "K").End(xlUp).Row
        If Not IsError(hoja.Cells(fila, 11).Value) Then
            Dim tipoPago As String
            tipoPago = LCase(hoja.Cells(fila, 11).Value)
            If tipoPago = "cheques" And UCase(hoja.Cells(fila, 3).Value) Like "*B*" Then
                suma = suma + hoja.Cells(fila, 5).Value
            End If
        End If
    Next fila
    SumarValoresCheques = suma
End Function


