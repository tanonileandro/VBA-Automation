VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("Historico Anual")
    Set rng = ws.Range("B:B")

    If Not Intersect(Target, rng) Is Nothing Then
        Application.EnableEvents = False

        For Each cell In Intersect(Target, rng)
            If cell.Value <> "" Then
                ws.Cells(cell.Row, 1).Value = Now
            End If
        Next cell

        Application.EnableEvents = True
    End If
End Sub

