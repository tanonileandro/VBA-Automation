Attribute VB_Name = "Filtro"
Sub FiltroDinamico()
'
' FiltroDinamico Macro
'

'
    Dim rngData As Range
    Dim rngCriteria As Range

    ' Definir el rango de datos y el rango de criterios
    Set rngData = Range("Tabla1[#All]")
    Set rngCriteria = Range("D2:H3")

    ' Verificar si hay resultados antes de aplicar el filtro
    If Application.WorksheetFunction.CountIf(rngData, rngCriteria.Cells(1)) = 0 Then
        MsgBox "No hay coincidencias para los criterios especificados.", vbInformation
        Exit Sub
    End If

    ' Filtrar los datos en el lugar
    rngData.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=rngCriteria, Unique:=False

    ' Verificar si se filtraron resultados
    If rngData.SpecialCells(xlCellTypeVisible).count = 1 Then
        MsgBox "No se encontraron resultados.", vbInformation
    End If
End Sub
Sub BorrarFiltro()
'
' BorrarFiltro Macro
'

'
    Range("D3:H3").Select
    Selection.ClearContents
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("D3").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("Tabla1[#All]").AdvancedFilter Action:=xlFilterInPlace, CriteriaRange _
        :=Range("D3:H3"), Unique:=False
End Sub

