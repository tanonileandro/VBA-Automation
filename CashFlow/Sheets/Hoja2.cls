VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
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
    valorA = Me.Range("A13").Value
    valorB = Me.Range("B13").Value
    
    ' Sumar 1 a valorA para la celda D25
    valorAModificado = valorA + 1
    valorAModificado2 = valorAModificado + 1
    
    ' Reiniciar las sumas
    sumaF = 0
    sumaH = 0
    
    ' Verificar si el cambio ocurrió en las celdas A13 o B13
    If Not Intersect(Target, Me.Range("A13:B13")) Is Nothing Then
        ' Buscar coincidencias y sumar los valores de la columna F y H
        For i = 2 To obligaciones.ListObjects("Tabla1").ListRows.Count + 1
            If obligaciones.Cells(i, 1).Value = valorA And obligaciones.Cells(i, 10).Value = valorB Then
                sumaF = sumaF + obligaciones.Cells(i, 6).Value
                sumaH = sumaH + obligaciones.Cells(i, 8).Value
            End If
        Next i
        
        ' Mensajes de depuración
        Debug.Print "sumaF: " & sumaF
        Debug.Print "sumaH: " & sumaH
        Debug.Print "valorAModificado: " & valorAModificado
        Debug.Print "valorAModificado2: " & valorAModificado2

        ' Poner el resultado de la suma de la columna F en la celda D20 con signo negativo
        Me.Range("D20").Value = -sumaF
        
        ' Poner el resultado de la suma de la columna H en la celda E20
        Me.Range("E20").Value = -sumaH
        
        ' Sumar valores en D25 y D32 considerando valorAModificado y valorAModificado2 y verificando también valorB
        Dim sumaD25 As Double
        Dim sumaD32 As Double
        Dim sumaE25 As Double
        Dim sumaE32 As Double

        sumaD25 = 0
        sumaD32 = 0
        sumaE25 = 0
        sumaE32 = 0

        For i = 2 To obligaciones.ListObjects("Tabla1").ListRows.Count + 1
            If obligaciones.Cells(i, 1).Value = valorAModificado And obligaciones.Cells(i, 10).Value = valorB Then
                sumaD25 = sumaD25 + obligaciones.Cells(i, 6).Value
                sumaE25 = sumaE25 + obligaciones.Cells(i, 8).Value
            End If
            If obligaciones.Cells(i, 1).Value = valorAModificado2 And obligaciones.Cells(i, 10).Value = valorB Then
                sumaD32 = sumaD32 + obligaciones.Cells(i, 6).Value
                sumaE32 = sumaE32 + obligaciones.Cells(i, 8).Value
            End If
        Next i

        Me.Range("D25").Value = -sumaD25
        Debug.Print "D25: " & Me.Range("D25").Value

        Me.Range("D32").Value = -sumaD32
        Debug.Print "D32: " & Me.Range("D32").Value

        Me.Range("E25").Value = -sumaE25
        Debug.Print "E25: " & Me.Range("E25").Value

        Me.Range("E32").Value = -sumaE32
        Debug.Print "E32: " & Me.Range("E32").Value
    End If
End Sub

