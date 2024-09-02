# ninamacros

Sub LimpiarDatos()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim sum As Double
    Dim count As Long
    
    ' Configura la hoja de trabajo activa
    Set ws = ActiveSheet
    
    ' Encuentra la última fila y columna utilizadas
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Eliminar filas completamente vacías
    For i = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i
    
    ' Recalcular la última fila después de eliminar filas vacías
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Completar celdas vacías con el promedio de la columna
    For i = 1 To lastCol
        Set rng = ws.Range(ws.Cells(1, i), ws.Cells(lastRow, i))
        sum = 0
        count = 0
        
        ' Calcular el promedio ignorando celdas vacías
        For Each cell In rng
            If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
                sum = sum + cell.Value
                count = count + 1
            End If
        Next cell
        
        ' Evitar la división por cero
        If count > 0 Then
            For Each cell In rng
                If IsEmpty(cell.Value) Then
                    cell.Value = sum / count
                End If
            Next cell
        End If
    Next i
    
    MsgBox "Limpieza de datos completada", vbInformation
End Sub
