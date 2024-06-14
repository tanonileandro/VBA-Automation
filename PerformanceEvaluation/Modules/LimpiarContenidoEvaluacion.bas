Attribute VB_Name = "LimpiarContenidoEvaluacion"
Sub LimpiarContenidoEvaluacion()
    Dim wsEvaluacion As Worksheet
    Dim lastRow As Long
    Dim answer As String

    ' Intenta establecer la hoja 'Evaluacion'
    On Error Resume Next
    Set wsEvaluacion = ThisWorkbook.Sheets("Evaluacion")
    On Error GoTo 0

    ' Comprueba si se encontró la hoja 'Evaluacion'
    If wsEvaluacion Is Nothing Then
        MsgBox "No se encontró la hoja 'Evaluacion'.", vbExclamation
        Exit Sub
    End If

    ' Encuentra la última fila en la columna A
    lastRow = wsEvaluacion.Cells(wsEvaluacion.Rows.count, "A").End(xlUp).Row

    ' Verifica si hay filas para limpiar
    If lastRow > 7 Then
        ' Pregunta al usuario si desea continuar
        answer = MsgBox("Atencion! Por favor Exporte los datos al 'Historico Anual' antes de vaciar la planilla 'Evaluacion'. ¿Desea continuar?", vbQuestion + vbYesNo)
        
        If answer = vbYes Then
            ' Limpia el contenido de las celdas en las columnas específicas desde la fila 7 hasta la última fila menos una
        wsEvaluacion.Range("L7:L" & lastRow - 1).ClearContents
        wsEvaluacion.Range("M7:M" & lastRow - 1).ClearContents
        wsEvaluacion.Range("N7:N" & lastRow - 1).ClearContents
        wsEvaluacion.Range("O7:O" & lastRow - 1).ClearContents
        wsEvaluacion.Range("P7:P" & lastRow - 1).ClearContents
        wsEvaluacion.Range("Q7:Q" & lastRow - 1).ClearContents
        wsEvaluacion.Range("R7:R" & lastRow - 1).ClearContents
        wsEvaluacion.Range("S7:S" & lastRow - 1).ClearContents
        wsEvaluacion.Range("T7:T" & lastRow - 1).ClearContents
        wsEvaluacion.Range("U7:U" & lastRow - 1).ClearContents
        wsEvaluacion.Range("V7:V" & lastRow - 1).ClearContents
        wsEvaluacion.Range("W7:W" & lastRow - 1).ClearContents
        wsEvaluacion.Range("X7:X" & lastRow - 1).ClearContents

            MsgBox "El contenido de las filas se limpió correctamente en la hoja 'Evaluacion'.", vbInformation
        Else
            MsgBox "La operación fue cancelada. No se eliminaron los datos.", vbInformation
        End If
    Else
        MsgBox "No hay filas con contenido para limpiar en la hoja 'Evaluacion'.", vbExclamation
    End If
End Sub


