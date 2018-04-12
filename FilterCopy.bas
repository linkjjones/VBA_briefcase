Attribute VB_Name = "FilterCopy"
Option Explicit
    Dim btmRow      As Long
    Dim rng         As Range
    Dim wsDump      As Worksheet
    Dim wsAll       As Worksheet
    Dim wsAPI510    As Worksheet
    Dim wsAPI570    As Worksheet
    Dim wsAPI653    As Worksheet
    Dim wsOther     As Worksheet
    Dim rngDump     As Range
    Dim FirstDataRow    As Long
    Dim lastCol         As Long
    Dim BT              As String
    Dim HighlightColor  As Long

Private Sub setStaticVariables()
    BT = "1-ENVIRONMENTAL"
    Set wsDump = Sheets("Data Dump")
    Set wsAll = Sheets("12 Mnth Outlook All Insp")
    Set wsAPI510 = Sheets("API 510")
    Set wsAPI570 = Sheets("API 570")
    Set wsAPI653 = Sheets("API 653")
    Set wsOther = Sheets("Other - Non API Insp")
    FirstDataRow = 3
    lastCol = Fn.GetLastCol(wsDump, 2, 10)
    HighlightColor = 36
End Sub

Public Sub ExtractData()
    On Error Resume Next
    
    setStaticVariables
    RemoveFilter wsDump
    
    'Get datadump range
    btmRow = Fn.GetLastRow(wsDump, 1, 2)
    lastCol = Fn.GetLastCol(wsDump, 2, 10)
    With wsDump
        Set rngDump = .Range(.Cells(2, 1), .Cells(btmRow, lastCol))
    End With
    
    'filter dump for bt
    rngDump.AutoFilter Field:=1, Criteria1:=BT
    
    'Remove filter wsAll
    RemoveFilter wsAll
    'clear 12 mnth outlook All Insp
    btmRow = Fn.GetLastRow(wsAll, 1, FirstDataRow)
    With wsAll
        'clear data
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
        rng.ClearContents
        rng.Interior.ColorIndex = xlNone
        Fn.ClearLines rng
    End With

    With wsDump.AutoFilter.Range
        Set rng = .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        If Err = 1004 Then 'no rows found
            'None
            wsAll.Cells(FirstDataRow, 1).Value = "None"
        Else
            rng.Copy wsAll.Cells(FirstDataRow, 1)
        End If
        On Error GoTo 0
        On Error Resume Next
    End With
    
    'format wsAll
    btmRow = Fn.GetLastRow(wsAll, 1, FirstDataRow)
    With wsAll
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    Fn.FormatRangeWithLines rng, True
    rng.Interior.ColorIndex = xlNone
    
    'add filter for API 510
    rngDump.AutoFilter Field:=4, Criteria1:=Array("EXCH", "FURN", _
                                 "PLBX", "PSAV", "PVSL"), Operator:=xlFilterValues
    'API510
    'remove filter
    RemoveFilter wsAPI510
    'clear
    btmRow = Fn.GetLastRow(wsAPI510, 1, FirstDataRow)
    With wsAPI510
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    'clear lines
    Fn.ClearLines rng
    'clear contents
    rng.ClearContents
    rng.Interior.ColorIndex = xlNone
    

    'copy/paste data
    With wsDump.AutoFilter.Range
        Set rng = .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        If Err = 1004 Then 'no rows found
            'None
            wsAPI510.Cells(FirstDataRow, 1).Value = "None"
        Else
            rng.Copy wsAPI510.Cells(FirstDataRow, 1)
        End If
        On Error GoTo 0
        On Error Resume Next
    End With
    
    'format wsAPI510
    btmRow = Fn.GetLastRow(wsAPI510, 1, FirstDataRow)
    With wsAPI510
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    Fn.FormatRangeWithLines rng, True
    rng.Interior.ColorIndex = xlNone
    
'API570
    'reset DataDump filter
    RemoveFilter wsDump
    rngDump.AutoFilter Field:=1, Criteria1:=BT
    rngDump.AutoFilter Field:=4, Criteria1:="PIPE", Operator:=xlFilterValues
    
    'clear wsAPI570
    RemoveFilter wsAPI570
    btmRow = Fn.GetLastRow(wsAPI570, 1, FirstDataRow)
    With wsAPI570
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    Fn.ClearLines rng
    rng.Interior.ColorIndex = xlNone
    rng.ClearContents
    
    'copy/paste data
'    With wsDump.AutoFilter.Range
'        Set rng = .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
'    End With
'    If Not rng Is Nothing Then
'        rng.Copy wsAPI570.Cells(FirstDataRow, 1)
'    End If
    With wsDump.AutoFilter.Range
        Set rng = .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        If Err = 1004 Then 'no rows found
            'None
            wsAPI570.Cells(FirstDataRow, 1).Value = "None"
        Else
            rng.Copy wsAPI570.Cells(FirstDataRow, 1)
        End If
        On Error GoTo 0
        On Error Resume Next
    End With
    
    
    'format wsAPI570
    btmRow = Fn.GetLastRow(wsAPI570, 1, FirstDataRow)
    With wsAPI570
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    Fn.FormatRangeWithLines rng, True
    rng.Interior.ColorIndex = xlNone
    
'API653
    'reset DataDump filter
    RemoveFilter wsDump
    rngDump.AutoFilter Field:=1, Criteria1:=BT
    rngDump.AutoFilter Field:=4, Criteria1:="TANK", Operator:=xlFilterValues
    
    'clear wsAPI653
    RemoveFilter wsAPI653
    btmRow = Fn.GetLastRow(wsAPI653, 1, FirstDataRow)
    With wsAPI653
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    Fn.ClearLines rng
    rng.Interior.ColorIndex = xlNone
    rng.ClearContents
    
    'copy/paste data
    With wsDump.AutoFilter.Range
        Set rng = .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        If Err = 1004 Then 'no rows found
            'None
            wsAPI653.Cells(FirstDataRow, 1).Value = "None"
        Else
            rng.Copy wsAPI653.Cells(FirstDataRow, 1)
        End If
        On Error GoTo 0
        On Error Resume Next
    End With
    
    'format wsAPI653
    btmRow = Fn.GetLastRow(wsAPI653, 1, FirstDataRow)
    With wsAPI653
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    Fn.FormatRangeWithLines rng, True
    rng.Interior.ColorIndex = xlNone
  
'Other
    'reset DataDump filter
    RemoveFilter wsDump
    rngDump.AutoFilter Field:=1, Criteria1:=BT
    rngDump.AutoFilter Field:=4, Criteria1:="MISC", Operator:=xlFilterValues
    
    'clear wsOther
    RemoveFilter wsOther
    btmRow = Fn.GetLastRow(wsOther, 1, FirstDataRow)
    With wsOther
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    Fn.ClearLines rng
    rng.Interior.ColorIndex = xlNone
    rng.ClearContents
    
    'copy/paste data
    With wsDump.AutoFilter.Range
        Set rng = .Offset(1, 0).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
        If Err = 1004 Then 'no rows found
            'None
            wsOther.Cells(FirstDataRow, 1).Value = "None"
        Else
            rng.Copy wsOther.Cells(FirstDataRow, 1)
        End If
        On Error GoTo 0
        On Error Resume Next
    End With
    
    'format wsOther
    btmRow = Fn.GetLastRow(wsOther, 1, FirstDataRow)
    With wsOther
        Set rng = .Range(.Cells(FirstDataRow, 1), .Cells(btmRow, lastCol))
    End With
    Fn.FormatRangeWithLines rng, True
    rng.Interior.ColorIndex = xlNone
    RemoveFilter wsDump
    
End Sub

Private Sub RemoveFilter(ws As Worksheet)
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
End Sub

Public Sub HighlightOverdueSheets()
    HighlightOverdue wsAll
    HighlightOverdue wsAPI510
    HighlightOverdue wsAPI570
    HighlightOverdue wsAPI653
    HighlightOverdue wsOther
End Sub

Private Sub HighlightOverdue(ws As Worksheet)
    Dim i               As Long
    Dim NextDateCol     As Long
    Dim currentMonth    As Integer
    setStaticVariables
    
    With ws
        'get data range
        btmRow = Fn.GetLastRow(ws, 1, FirstDataRow)
        'LastCol
        NextDateCol = Fn.HeaderCol(ws, "Next Date", 2)
        For i = FirstDataRow To btmRow
            If Not .Cells(i, NextDateCol) = "" Then
                If CDate(.Cells(i, NextDateCol)) < Date And _
                Format(Date, "m-yyyy") <> Format(CDate(.Cells(i, NextDateCol)), "m-yyyy") Then
                    .Range(.Cells(i, 1), .Cells(i, lastCol)).Interior.ColorIndex = HighlightColor
                End If
            End If
        Next i
        
    End With
End Sub
