Attribute VB_Name = "Fn"
Option Explicit

Public Function GetLastRow(ws As Worksheet, ColumnNumber As Long, _
                           StartRow As Long, _
                           Optional ToColumn As Long, _
                           Optional Contiguous As Boolean) As Long
    Dim lastRow     As Long
    Dim i As Long, j As Long
    Dim ColLastRow  As Long
    
    If Contiguous Then
        'This function simply loops down until it finds a blank and returns the last
        'filled cell row.
        'If there are no rows (not counting the hearder row, then return 0
        
        If ToColumn = 0 Then
            ToColumn = ColumnNumber
        End If
        
        i = StartRow
        For j = ColumnNumber To ToColumn
            Do Until ws.Cells(i + 1, j) = ""
                i = i + 1
            Loop
        Next j
        
        GetLastRow = i
    Else
        ' xl'ing up
        If ToColumn = 0 Then
            ToColumn = ColumnNumber
        End If
        
        If ColumnNumber = 0 Then
            ColumnNumber = 1
        End If
            
        'loop through columns and get the greatest last row
        For i = ColumnNumber To ToColumn
            ColLastRow = ws.Cells(ws.Rows.Count, i).End(xlUp).Row
            lastRow = IIf(ColLastRow > lastRow, ColLastRow, lastRow)
        Next i
        
        GetLastRow = IIf(lastRow < StartRow, StartRow, lastRow)
    End If
    
End Function

Public Sub FormatRangeWithLines(FormatRange As Range, Optional VerticalLines As Boolean)
    
    With FormatRange
        If VerticalLines Then
            .Borders(xlInsideHorizontal).Color = RGB(0, 0, 0)
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlInsideVertical).Color = RGB(0, 0, 0)
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        Else
            .Borders(xlInsideHorizontal).Color = RGB(200, 200, 200)
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlEdgeBottom).Color = RGB(100, 100, 100)
        End If
    End With
    
End Sub

Public Sub ClearLines(rng As Range)
    rng.Borders.LineStyle = xlNone
    rng.Borders.ColorIndex = xlNone
End Sub

Public Function GetLastCol(ws As Worksheet, RowNumber As Long, _
                           Optional ColLimit As Long) As Long

    GetLastCol = ws.Cells(RowNumber, ws.Columns.Count).End(xlToLeft).Column
    
    GetLastCol = IIf(GetLastCol < ColLimit, ColLimit, GetLastCol)
End Function

Public Function HeaderCol(ws As Worksheet, HeaderName As String, HeadingRow As Long, _
                          Optional LastOccurance As Boolean) As Long
    On Error Resume Next
    Dim Header As Range
    
    Dim LookRange As Range, cell As Range
    Dim LastDataColumn As Long
    
'    HeadingRow = IIf(HeadingRow = 0, HeaderRow, HeadingRow)
    
'    LastDataColumn = GetLastCol(ws, HeadingRow)
    LastDataColumn = ws.UsedRange.Columns.Count
    With ws
        'since the above code is flaky...lets just loop through
        Set LookRange = .Range(.Cells(HeadingRow, 1), .Cells(HeadingRow, LastDataColumn))
        For Each cell In LookRange
            If cell.Value = HeaderName Then
                If Not LastOccurance Then
                    HeaderCol = cell.Column
                    Exit For
                Else
                    HeaderCol = cell.Column
                End If
            End If
        Next cell
        
    End With

End Function
