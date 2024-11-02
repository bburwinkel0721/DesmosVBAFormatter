Attribute VB_Name = "Module1"
Sub CalculateScores()

    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long

    Set ws = ActiveSheet ' Apply to the active sheet

    ' Clear the second and third columns
    ws.Columns(2).ClearContents
    ws.Columns(3).ClearContents

    ' Find the last column
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Delete the last column
    ws.Columns(lastCol).Delete
    lastCol = lastCol - 1 ' Adjust lastCol to reflect the new last column

    ' Label the headers
    ws.Cells(1, 2).Value = "Raw Score"
    ws.Cells(1, 3).Value = "Percentage"

    ' Find the last row with data in column 1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Replace cells containing only "-" with 0
    For i = 1 To lastRow
        For j = 1 To lastCol
            If ws.Cells(i, j).Value = "-" Then
                ws.Cells(i, j).Value = 0
            End If
        Next j
    Next i

    ' Set formulas in each row below the headers
    For i = 2 To lastRow
        ' Set formula for "Raw Score" in column 2
        ws.Cells(i, 2).Formula = "=4 * " & ws.Cells(i, 3).Address

        ' Set formula for "Percentage" in column 3
        ws.Cells(i, 3).Formula = "=SUM(" & ws.Cells(i, 4).Address & ":" & ws.Cells(i, lastCol).Address & ") / (COUNT(" & ws.Cells(i, 4).Address & ":" & ws.Cells(i, lastCol).Address & ") * 4)"
    Next i

    ' Add a row for averages
    ws.Cells(lastRow + 1, 1).Value = "Average" ' Label for the average row
    For j = 2 To lastCol
        ws.Cells(lastRow + 1, j).Formula = "=AVERAGE(" & ws.Cells(2, j).Address & ":" & ws.Cells(lastRow, j).Address & ")"
    Next j

End Sub
