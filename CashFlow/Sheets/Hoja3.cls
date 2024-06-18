VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim obligaciones As Worksheet
    Dim valorA As String
    Dim valorAModificado As String
    Dim valorAModificado2 As String
    Dim valorB As String
    Dim sumaF As Double
    Dim sumaH As Double
    Dim i As Long
    
    ' Definir la hoja de trabajo de las obligaciones
    Set obligaciones = ThisWorkbook.Sheets("Obligaciones")
    
    ' Obtener los valores de las celdas A13 y B13
    valorA = Me.Range("B14").Value
    valorB = Me.Range("C14").Value
    
    sumaF = 0
    sumaH = 0
    ' Verificar si el cambio ocurrió en las celdas A13 o B13
    If Not Intersect(Target, Me.Range("B14:C14")) Is Nothing Then
        ' Buscar coincidencias y sumar los valores de la columna F y H
        For i = 2 To obligaciones.ListObjects("Tabla1").ListRows.Count + 1
            If obligaciones.Cells(i, 3).Value = valorA And obligaciones.Cells(i, 10).Value = valorB Then
                sumaF = sumaF + obligaciones.Cells(i, 6).Value
                sumaH = sumaH + obligaciones.Cells(i, 8).Value
            End If
        Next i

        ' Poner el resultado de la suma de la columna F en la celda D20 con signo negativo
        Me.Range("E21").Value = -sumaF
        
        ' Poner el resultado de la suma de la columna H en la celda E20
        Me.Range("F21").Value = -sumaH
    End If
End Sub
