Attribute VB_Name = "OcultarMostrar"
Sub AlternarColumnas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim columnas As Variant
    columnas = Array("D", "F")

    Dim col As Variant
    For Each col In columnas
        If ws.Columns(col).Hidden Then
            ws.Columns(col).Hidden = False
        Else
            ws.Columns(col).Hidden = True
        End If
    Next col
End Sub

