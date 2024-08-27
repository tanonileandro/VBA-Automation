VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub OpenFormCarga_Click()
    FormImpNuevo.Show
End Sub

Sub OpenFormPago_Click()
    FormImpPago.Show
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range
    Dim colFecha As ListColumn
    
    ' Intentar obtener la columna de fecha por su nombre
    Set colFecha = Me.ListObjects(1).ListColumns("Fecha Vto")
    
    ' Salir si no se encontró la columna
    If colFecha Is Nothing Then Exit Sub
    
    ' Verificar si el cambio ocurrió en la columna de fecha
    Set rng = Intersect(Target, colFecha.DataBodyRange)
    If Not rng Is Nothing Then
        ' Llamar al procedimiento para actualizar las semanas
        ActualizarSemanas
    End If
End Sub



