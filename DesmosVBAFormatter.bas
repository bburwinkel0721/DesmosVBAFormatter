Attribute VB_Name = "Module1"
Sub CalculateScores()

    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    
    Set ws = ActiveSheet ' Apply to the active sheet

    ' Clear the second and third columns and first cell
    ws.Columns(2).ClearContents
    ws.Columns(3).ClearContents
    ws.Cells(1, 1).ClearContents

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
    
    ' Adds an additional row below the header row
    ws.Rows(2).Insert Shift:=xlDown
    
    ' Place raw variable and Label
    ws.Cells(2, 2).Value = "=4"
    ws.Cells(2, 1).Value = "Max Raw Score"

    ' Set formulas in each row below the headers
    For i = 3 To lastRow
        ' Set formula for "Raw Score" in column 2
        ws.Cells(i, 2).Formula = "=" & ws.Cells(2, 2).Address & "*" & ws.Cells(i, 3).Address

        ' Set formula for "Percentage" in column 3
        ws.Cells(i, 3).Formula = "=SUM(" & ws.Cells(i, 4).Address & ":" & ws.Cells(i, lastCol).Address & ") / (COUNT(" & ws.Cells(i, 4).Address & ":" & ws.Cells(i, lastCol).Address & ") * 4)"
    Next i

    ' Add a row for averages
    ws.Cells(lastRow + 1, 1).Value = "Mean" ' Label for the average row
    For j = 2 To lastCol
        ws.Cells(lastRow + 1, j).Formula = "=AVERAGE(" & ws.Cells(2, j).Address & ":" & ws.Cells(lastRow, j).Address & ")"
    Next j
    
    'This next section applies our number formatting and resizing of cells
    Columns("B:B").NumberFormat = "0.00"
    Rows(lastRow + 1).NumberFormat = "0.00"
    Columns("C:C").NumberFormat = "0.00%"
    Columns("A:Z").EntireColumn.AutoFit
    
    'This section applies color formatting and styling to the table
    
    '
    With ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
        .Borders.LineStyle = xlContinuous    ' Apply borders to all sides
        .Borders.Weight = xlThin             ' Set the thickness of the borders
    End With
    
    ' Adds alternating color to the main section of the table
    For i = 2 To lastRow ' Start from row 2 if row 1 is a header
        If i Mod 2 = 0 Then
            ' Apply light red fill for even rows
            ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Interior.Color = RGB(255, 204, 204) ' Light red
        Else
            ' Apply white fill for odd rows
            ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Interior.Color = RGB(255, 255, 255) ' White
        End If
    Next i
    
    ' This section adds conditional formatting for the grades for better grade analysis
    For Each cell In ws.Range("C3:C" & lastRow)
        If IsNumeric(cell.Value) Then ' Check if the cell contains a number
            Select Case cell.Value
                Case Is >= 0.9
                    cell.Interior.Color = RGB(173, 216, 230) ' Light Blue
                Case Is >= 0.8
                    cell.Interior.Color = RGB(144, 238, 144) ' Light Green
                Case Is >= 0.7
                    cell.Interior.Color = RGB(255, 255, 153) ' Light Yellow
                Case Is >= 0.6
                    cell.Interior.Color = RGB(255, 204, 153) ' Light Orange
                Case Else
                    cell.Interior.Color = RGB(255, 153, 153) ' Light Red
            End Select
        End If
    Next cell
    
    'Adjusts the font and sizeto be the same through out the spreadsheet
    With ws.Cells
        .Font.Name = "Calibri" ' Set the font name (change to desired font)
        .Font.Size = 12        ' Set the font size (change to desired size)
    End With
    
    'Formatting for the header
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Interior.Color = RGB(255, 153, 153) ' Darker Red
        .Borders.LineStyle = xlContinuous    ' Apply borders to all sides
        .Borders.Weight = xlThin             ' Set the thickness of the borders
    End With
    
    'Formatting for the last row
    With ws.Range(ws.Cells(lastRow + 1, 1), ws.Cells(lastRow + 1, lastCol))
        .Interior.Color = RGB(255, 153, 153) ' Darker Red
        .Borders.LineStyle = xlContinuous    ' Apply borders to all sides
        .Borders.Weight = xlThin             ' Set the thickness of the borders
    End With
    
    ' Center align all numeric values in the worksheet
    With ws.Range(ws.Cells(2, 2), ws.Cells(lastRow + 1, lastCol))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With


End Sub
