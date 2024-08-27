Attribute VB_Name = "DepurarYActSemana"
Sub DepurarTablaYActSem()

    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim numeroSemana As Integer
    Dim Hoja As Worksheet
    Dim i As Long
    Dim tabla As ListObject
    Dim columnaLIndex As Integer
    
    ' Establecer la hoja de trabajo activa
    Set Hoja = ActiveSheet
    
    ' Calcular la fecha de inicio de la semana (lunes)
    fechaInicio = Date - (Weekday(Date, vbMonday) - 1)
    
    ' Calcular la fecha de fin de la semana (domingo)
    fechaFin = fechaInicio + 6
    
    ' Calcular el número de semana
    numeroSemana = Application.WorksheetFunction.WeekNum(Date, vbMonday)
    
    ' Escribir las fechas y el número de semana en las celdas A1 y A2
    With Hoja
        .Range("A1").Value = "Semana " & numeroSemana
        .Range("A2").Value = "Semana " & Format(fechaInicio, "dd-mm") & " al " & Format(fechaFin, "dd-mm")
    End With
    
    ' Asumir que hay solo una tabla en la hoja activa
    On Error Resume Next
    Set tabla = Hoja.ListObjects(1)
    On Error GoTo 0
    
    If tabla Is Nothing Then
        MsgBox "No se encontró ninguna tabla en la hoja activa.", vbExclamation
        Exit Sub
    End If
    
    ' Suponiendo que la columna L corresponde a la columna 12 dentro de la tabla
    columnaLIndex = 12
    
    ' Verificar si la columna L está dentro del rango de columnas de la tabla
    If columnaLIndex > tabla.ListColumns.Count Then
        MsgBox "La tabla en la hoja activa no tiene suficientes columnas para incluir la columna L.", vbExclamation
        Exit Sub
    End If
    
    ' Eliminar filas donde el valor en la columna L sea mayor que numeroSemana
    For i = tabla.ListRows.Count To 1 Step -1
        If IsNumeric(tabla.DataBodyRange(i, columnaLIndex).Value) Then
            If tabla.DataBodyRange(i, columnaLIndex).Value > numeroSemana Then
                tabla.ListRows(i).Delete
            End If
        End If
    Next i
    
End Sub






