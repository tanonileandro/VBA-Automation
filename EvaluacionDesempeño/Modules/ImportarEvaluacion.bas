Attribute VB_Name = "ImportarEvaluacion"
Sub ExtraerDatosDM()

    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    
    Dim wsEval As Worksheet
    Dim wsDM As Worksheet
    Dim tblEval As ListObject
    Dim rngEval As Range
    Dim rngDM As Range
    Dim cell As Range
    Dim foundCell As Range
    Dim wbEval As Workbook
    Dim wbIsOpened As Boolean
    
    On Error GoTo ErrorHandler
    
    Set wsEval = ThisWorkbook.Sheets("Evaluacion")
    
    ' Verificar si el archivo está abierto por otro usuario
    On Error Resume Next
    Set wbEval = Workbooks("Evaluacion de Desempeño-Dario Muñoz.xlsm")
    On Error GoTo ErrorHandler
    
    If wbEval Is Nothing Then
        ' Abrir la planilla si está cerrada
        Set wbEval = Workbooks.Open("Z:\ADMINISTRACION\RRHH\RRHH\EVALUACIONES DE DESEMPEÑO\Evaluacion de Desempeño-Dario Muñoz.xlsm")
        wbIsOpened = False
    Else
        ' Guardar el archivo si está abierto por otro usuario
        If wbEval.ReadOnly Then
            wbEval.Save
        End If
        wbIsOpened = True
    End If
    
    ' Abrir la hoja "Mto" del archivo "Evaluacion de Desempeño-Mantenimiento.xlsm"
    Set wsDM = wbEval.Sheets("DM")
    
    ' Referencia a la tabla "Tabla1" en la hoja "Evaluacion"
    Set tblEval = wsEval.ListObjects("Tabla1")
    ' Rango de datos de la tabla "Tabla1"
    Set rngEval = tblEval.DataBodyRange
    
    For Each cell In rngEval.Columns(1).Cells ' Considerando solo la primera columna de la tabla
        Set foundCell = wsDM.Columns("A").Find(cell.Value, LookIn:=xlValues, lookat:=xlWhole)
        If Not foundCell Is Nothing Then
            wsEval.Cells(cell.Row, "L").Value = foundCell.Offset(0, 4).Value ' Columna D
            wsEval.Cells(cell.Row, "M").Value = foundCell.Offset(0, 5).Value ' Columna E
            wsEval.Cells(cell.Row, "N").Value = foundCell.Offset(0, 6).Value ' Columna F
            wsEval.Cells(cell.Row, "O").Value = foundCell.Offset(0, 7).Value ' Columna G
            wsEval.Cells(cell.Row, "P").Value = foundCell.Offset(0, 8).Value ' Columna H
            wsEval.Cells(cell.Row, "Q").Value = foundCell.Offset(0, 9).Value ' Columna I
            wsEval.Cells(cell.Row, "R").Value = foundCell.Offset(0, 10).Value ' Columna J
            wsEval.Cells(cell.Row, "S").Value = foundCell.Offset(0, 11).Value ' Columna K
        End If
    Next cell
    
    If Not wbIsOpened Then
        wbEval.Close SaveChanges:=False
    End If
    
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: No se pudo encontrar la planilla o se produjo un error al buscar los datos.", vbExclamation

End Sub

Sub ExtraerDatosMA()

    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    
    Dim wsEval As Worksheet
    Dim wsDM As Worksheet
    Dim tblEval As ListObject
    Dim rngEval As Range
    Dim rngDM As Range
    Dim cell As Range
    Dim foundCell As Range
    Dim wbEval As Workbook
    Dim wbIsOpened As Boolean
    
    On Error GoTo ErrorHandler
    
    Set wsEval = ThisWorkbook.Sheets("Evaluacion")
    
    ' Verificar si el archivo está abierto por otro usuario
    On Error Resume Next
    Set wbEval = Workbooks("Evaluacion de Desempeño-Matias Arriola.xlsm")
    On Error GoTo ErrorHandler
    
    If wbEval Is Nothing Then
        ' Abrir la planilla si está cerrada
        Set wbEval = Workbooks.Open("Z:\ADMINISTRACION\RRHH\RRHH\EVALUACIONES DE DESEMPEÑO\Evaluacion de Desempeño-Matias Arriola.xlsm")
        wbIsOpened = False
    Else
        ' Guardar el archivo si está abierto por otro usuario
        If wbEval.ReadOnly Then
            wbEval.Save
        End If
        wbIsOpened = True
    End If
    
    ' Abrir la hoja "MA" del archivo "Evaluacion de Desempeño-Matias Arriola.xlsm"
    Set wsDM = wbEval.Sheets("MA")
    
    ' Referencia a la tabla "Tabla1" en la hoja "Evaluacion"
    Set tblEval = wsEval.ListObjects("Tabla1")
    ' Rango de datos de la tabla "Tabla1"
    Set rngEval = tblEval.DataBodyRange
    
    For Each cell In rngEval.Columns(1).Cells ' Considerando solo la primera columna de la tabla
        Set foundCell = wsDM.Columns("A").Find(cell.Value, LookIn:=xlValues, lookat:=xlWhole)
        If Not foundCell Is Nothing Then
            wsEval.Cells(cell.Row, "L").Value = foundCell.Offset(0, 4).Value ' Columna D
            wsEval.Cells(cell.Row, "M").Value = foundCell.Offset(0, 5).Value ' Columna E
            wsEval.Cells(cell.Row, "N").Value = foundCell.Offset(0, 6).Value ' Columna F
            wsEval.Cells(cell.Row, "O").Value = foundCell.Offset(0, 7).Value ' Columna G
            wsEval.Cells(cell.Row, "P").Value = foundCell.Offset(0, 8).Value ' Columna H
            wsEval.Cells(cell.Row, "Q").Value = foundCell.Offset(0, 9).Value ' Columna I
            wsEval.Cells(cell.Row, "R").Value = foundCell.Offset(0, 10).Value ' Columna J
            wsEval.Cells(cell.Row, "S").Value = foundCell.Offset(0, 11).Value ' Columna K
        End If
    Next cell
    
    ' Cerrar el archivo "Evaluacion de Desempeño-Matias Arriola.xlsm" si fue abierto en esta ejecución
    If Not wbIsOpened Then
        wbEval.Close SaveChanges:=False
    End If
    
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: No se pudo encontrar la planilla o se produjo un error al buscar los datos.", vbExclamation

End Sub

Sub ExtraerDatosVB()

    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    
    Dim wsEval As Worksheet
    Dim wsDM As Worksheet
    Dim tblEval As ListObject
    Dim rngEval As Range
    Dim rngDM As Range
    Dim cell As Range
    Dim foundCell As Range
    Dim wbEval As Workbook
    Dim wbIsOpened As Boolean
    
    On Error GoTo ErrorHandler
    
    Set wsEval = ThisWorkbook.Sheets("Evaluacion")
    
    ' Verificar si el archivo está abierto por otro usuario
    On Error Resume Next
    Set wbEval = Workbooks("Evaluacion de Desempeño-Victor Benavidez.xlsm")
    On Error GoTo ErrorHandler
    
    If wbEval Is Nothing Then
        ' Abrir la planilla si está cerrada
        Set wbEval = Workbooks.Open("Z:\ADMINISTRACION\RRHH\RRHH\EVALUACIONES DE DESEMPEÑO\Evaluacion de Desempeño-Victor Benavidez.xlsm")
        wbIsOpened = False
    Else
        ' Guardar el archivo si está abierto por otro usuario
        If wbEval.ReadOnly Then
            wbEval.Save
        End If
        wbIsOpened = True
    End If
    
    ' Abrir la hoja "VB" del archivo "Evaluacion de Desempeño-Victor Benavidez.xlsm"
    Set wsDM = wbEval.Sheets("VB")
    
    ' Referencia a la tabla "Tabla1" en la hoja "Evaluacion"
    Set tblEval = wsEval.ListObjects("Tabla1")
    ' Rango de datos de la tabla "Tabla1"
    Set rngEval = tblEval.DataBodyRange
    
    For Each cell In rngEval.Columns(1).Cells ' Considerando solo la primera columna de la tabla
        Set foundCell = wsDM.Columns("A").Find(cell.Value, LookIn:=xlValues, lookat:=xlWhole)
        If Not foundCell Is Nothing Then
            wsEval.Cells(cell.Row, "L").Value = foundCell.Offset(0, 4).Value ' Columna D
            wsEval.Cells(cell.Row, "M").Value = foundCell.Offset(0, 5).Value ' Columna E
            wsEval.Cells(cell.Row, "N").Value = foundCell.Offset(0, 6).Value ' Columna F
            wsEval.Cells(cell.Row, "O").Value = foundCell.Offset(0, 7).Value ' Columna G
            wsEval.Cells(cell.Row, "P").Value = foundCell.Offset(0, 8).Value ' Columna H
            wsEval.Cells(cell.Row, "Q").Value = foundCell.Offset(0, 9).Value ' Columna I
            wsEval.Cells(cell.Row, "R").Value = foundCell.Offset(0, 10).Value ' Columna J
            wsEval.Cells(cell.Row, "S").Value = foundCell.Offset(0, 11).Value ' Columna K
        End If
    Next cell
    
    ' Cerrar el archivo "Evaluacion de Desempeño-Victor Benavidez.xlsm" si fue abierto en esta ejecución
    If Not wbIsOpened Then
        wbEval.Close SaveChanges:=False
    End If
    
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: No se pudo encontrar la planilla o se produjo un error al buscar los datos.", vbExclamation

End Sub

Sub ExtraerDatosMto()

    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    
    Dim wsEval As Worksheet
    Dim wsDM As Worksheet
    Dim tblEval As ListObject
    Dim rngEval As Range
    Dim rngDM As Range
    Dim cell As Range
    Dim foundCell As Range
    Dim wbEval As Workbook
    Dim wbIsOpened As Boolean
    
    On Error GoTo ErrorHandler
    
    Set wsEval = ThisWorkbook.Sheets("Evaluacion")
    
    ' Verificar si el archivo está abierto por otro usuario
    On Error Resume Next
    Set wbEval = Workbooks("Evaluacion de Desempeño-Mantenimiento.xlsm")
    On Error GoTo ErrorHandler
    
    If wbEval Is Nothing Then
        ' Abrir la planilla si está cerrada
        Set wbEval = Workbooks.Open("Z:\ADMINISTRACION\RRHH\RRHH\EVALUACIONES DE DESEMPEÑO\Evaluacion de Desempeño-Mantenimiento.xlsm")
        wbIsOpened = False
    Else
        ' Guardar el archivo si está abierto por otro usuario
        If wbEval.ReadOnly Then
            wbEval.Save
        End If
        wbIsOpened = True
    End If
    
    ' Abrir la hoja "Mto" del archivo "Evaluacion de Desempeño-Mantenimiento.xlsm"
    Set wsDM = wbEval.Sheets("Mto")
    
    ' Referencia a la tabla "Tabla1" en la hoja "Evaluacion"
    Set tblEval = wsEval.ListObjects("Tabla1")
    ' Rango de datos de la tabla "Tabla1"
    Set rngEval = tblEval.DataBodyRange
    
    For Each cell In rngEval.Columns(1).Cells ' Considerando solo la primera columna de la tabla
        Set foundCell = wsDM.Columns("A").Find(cell.Value, LookIn:=xlValues, lookat:=xlWhole)
        If Not foundCell Is Nothing Then
            wsEval.Cells(cell.Row, "L").Value = foundCell.Offset(0, 4).Value ' Columna D
            wsEval.Cells(cell.Row, "M").Value = foundCell.Offset(0, 5).Value ' Columna E
            wsEval.Cells(cell.Row, "N").Value = foundCell.Offset(0, 6).Value ' Columna F
            wsEval.Cells(cell.Row, "O").Value = foundCell.Offset(0, 7).Value ' Columna G
            wsEval.Cells(cell.Row, "P").Value = foundCell.Offset(0, 8).Value ' Columna H
            wsEval.Cells(cell.Row, "Q").Value = foundCell.Offset(0, 9).Value ' Columna I
            wsEval.Cells(cell.Row, "R").Value = foundCell.Offset(0, 10).Value ' Columna J
            wsEval.Cells(cell.Row, "S").Value = foundCell.Offset(0, 11).Value ' Columna K
        End If
    Next cell
    
    ' Cerrar el archivo "Evaluacion de Desempeño-Mantenimiento.xlsm" si fue abierto en esta ejecución
    If Not wbIsOpened Then
        wbEval.Close SaveChanges:=False
    End If
    
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: No se pudo encontrar la planilla o se produjo un error al buscar los datos.", vbExclamation

End Sub

Sub ExtraerDatosLL()

    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    
    Dim wsEval As Worksheet
    Dim wsDM As Worksheet
    Dim tblEval As ListObject
    Dim rngEval As Range
    Dim rngDM As Range
    Dim cell As Range
    Dim foundCell As Range
    Dim wbEval As Workbook
    Dim wbIsOpened As Boolean
    
    On Error GoTo ErrorHandler
    
    Set wsEval = ThisWorkbook.Sheets("Evaluacion")
    
    ' Verificar si el archivo está abierto por otro usuario
    On Error Resume Next
    Set wbEval = Workbooks("Evaluacion de Desempeño-Lorena Lucena.xlsm")
    On Error GoTo ErrorHandler
    
    If wbEval Is Nothing Then
        ' Abrir la planilla si está cerrada
        Set wbEval = Workbooks.Open("Z:\ADMINISTRACION\RRHH\RRHH\EVALUACIONES DE DESEMPEÑO\Evaluacion de Desempeño-Lorena Lucena.xlsm")
        wbIsOpened = False
    Else
        ' Guardar el archivo si está abierto por otro usuario
        If wbEval.ReadOnly Then
            wbEval.Save
        End If
        wbIsOpened = True
    End If
    
    ' Abrir la hoja "LL" del archivo "Evaluacion de Desempeño-Lorena Lucena.xlsm"
    Set wsDM = wbEval.Sheets("LL")
    
    ' Referencia a la tabla "Tabla1" en la hoja "Evaluacion"
    Set tblEval = wsEval.ListObjects("Tabla1")
    ' Rango de datos de la tabla "Tabla1"
    Set rngEval = tblEval.DataBodyRange
    
    For Each cell In rngEval.Columns(1).Cells ' Considerando solo la primera columna de la tabla
        ' Buscar el valor en la columna A de la hoja "LL"
        Set foundCell = wsDM.Columns("A").Find(cell.Value, LookIn:=xlValues, lookat:=xlWhole)
        If Not foundCell Is Nothing Then
            wsEval.Cells(cell.Row, "L").Value = foundCell.Offset(0, 4).Value ' Columna D
            wsEval.Cells(cell.Row, "M").Value = foundCell.Offset(0, 5).Value ' Columna E
            wsEval.Cells(cell.Row, "N").Value = foundCell.Offset(0, 6).Value ' Columna F
            wsEval.Cells(cell.Row, "O").Value = foundCell.Offset(0, 7).Value ' Columna G
            wsEval.Cells(cell.Row, "P").Value = foundCell.Offset(0, 8).Value ' Columna H
            wsEval.Cells(cell.Row, "Q").Value = foundCell.Offset(0, 9).Value ' Columna I
            wsEval.Cells(cell.Row, "R").Value = foundCell.Offset(0, 10).Value ' Columna J
            wsEval.Cells(cell.Row, "S").Value = foundCell.Offset(0, 11).Value ' Columna K
        End If
    Next cell
    
    ' Cerrar el archivo "Evaluacion de Desempeño-Lorena Lucena.xlsm" si fue abierto en esta ejecución
    If Not wbIsOpened Then
        wbEval.Close SaveChanges:=False
    End If
    
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: No se pudo encontrar la planilla o se produjo un error al buscar los datos.", vbExclamation

End Sub
Sub ExtraerDatosLR()

    Application.AskToUpdateLinks = False
    Application.ScreenUpdating = False
    
    Dim wsEval As Worksheet
    Dim wsDM As Worksheet
    Dim tblEval As ListObject
    Dim rngEval As Range
    Dim rngDM As Range
    Dim cell As Range
    Dim foundCell As Range
    Dim wbEval As Workbook
    Dim wbIsOpened As Boolean
    
    On Error GoTo ErrorHandler
    
    Set wsEval = ThisWorkbook.Sheets("Evaluacion")
    
    ' Verificar si el archivo está abierto por otro usuario
    On Error Resume Next
    Set wbEval = Workbooks("Evaluacion de Desempeño-Leonardo Rodriguez.xlsm")
    On Error GoTo ErrorHandler
    
    If wbEval Is Nothing Then
        ' Abrir la planilla si está cerrada
        Set wbEval = Workbooks.Open("Z:\ADMINISTRACION\RRHH\RRHH\EVALUACIONES DE DESEMPEÑO\Evaluacion de Desempeño-Leonardo Rodriguez.xlsm")
        wbIsOpened = False
    Else
        ' Guardar el archivo si está abierto por otro usuario
        If wbEval.ReadOnly Then
            wbEval.Save
        End If
        wbIsOpened = True
    End If
    
    ' Abrir la hoja "LR" del archivo "Evaluacion de Desempeño-Leonardo Rodriguez.xlsm"
    Set wsDM = wbEval.Sheets("LR")
    
    ' Referencia a la tabla "Tabla1" en la hoja "Evaluacion"
    Set tblEval = wsEval.ListObjects("Tabla1")
    ' Rango de datos de la tabla "Tabla1"
    Set rngEval = tblEval.DataBodyRange
    
    For Each cell In rngEval.Columns(1).Cells ' Considerando solo la primera columna de la tabla
        ' Buscar el valor en la columna A de la hoja "DM"
        Set foundCell = wsDM.Columns("A").Find(cell.Value, LookIn:=xlValues, lookat:=xlWhole)
        If Not foundCell Is Nothing Then
            wsEval.Cells(cell.Row, "L").Value = foundCell.Offset(0, 4).Value ' Columna D
            wsEval.Cells(cell.Row, "M").Value = foundCell.Offset(0, 5).Value ' Columna E
            wsEval.Cells(cell.Row, "N").Value = foundCell.Offset(0, 6).Value ' Columna F
            wsEval.Cells(cell.Row, "O").Value = foundCell.Offset(0, 7).Value ' Columna G
            wsEval.Cells(cell.Row, "P").Value = foundCell.Offset(0, 8).Value ' Columna H
            wsEval.Cells(cell.Row, "Q").Value = foundCell.Offset(0, 9).Value ' Columna I
            wsEval.Cells(cell.Row, "R").Value = foundCell.Offset(0, 10).Value ' Columna J
            wsEval.Cells(cell.Row, "S").Value = foundCell.Offset(0, 11).Value ' Columna K
        End If
    Next cell
    
    ' Cerrar el archivo "Evaluacion de Desempeño-Leonardo Rodriguez.xlsm" si fue abierto en esta ejecución
    If Not wbIsOpened Then
        wbEval.Close SaveChanges:=False
    End If
    
    Application.AskToUpdateLinks = True
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    MsgBox "Error: No se pudo encontrar la planilla o se produjo un error al buscar los datos.", vbExclamation

End Sub



