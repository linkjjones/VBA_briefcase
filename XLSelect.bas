Attribute VB_Name = "XLSelect"
Option Explicit

Public Sub SelectionFormat(ws As Worksheet, Target As Range, rng As Range, _
                     UpdateText As String, _
                     Optional ValidationColOffset As Long)
    'This routine finds the intersection and formats
    Dim IntersectRange As Range
    
    Application.EnableEvents = False
    
    Set IntersectRange = Application.Intersect(rng, Target)
    If IntersectRange Is Nothing Then
        Exit Sub
    End If
    
    With IntersectRange
        If ValidationColOffset = 0 Or _
           IntersectRange.Cells(1).Offset(, ValidationColOffset).Value <> "" Then
            .Value = UpdateText
            If .Cells(1).Value = "Trash" Or _
                .Cells(1).Value = "Delete" Or _
                .Cells(1).Value = "X" Then
                .Interior.Color = vbBlack
                .Font.Color = RGB(255, 0, 0)
'                    .Borders.Color = RGB(255, 0, 0)
            ElseIf .Cells(1).Value = "Update" Or _
                .Cells(1).Value = ChrW(&H2713) Then
                .Interior.ColorIndex = Orange
                .Font.Color = RGB(255, 255, 255)
'                    .Borders.ColorIndex = xlNone
            ElseIf .Cells(1).Value = "Restore" Then
                .Interior.Color = RGB(0, 153, 0)
                .Font.Color = vbWhite
'                    .Borders.ColorIndex = xlNone
            Else
                .Value = ""
                .Interior.ColorIndex = 0
                .Borders.ColorIndex = xlNone
            End If
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Application.EnableEvents = True
    
End Sub

Public Function SelectionCycle(CurrentValue As String) As String
    
    If CurrentValue = "" Then
        SelectionCycle = ChrW(&H2713)
    ElseIf CurrentValue = ChrW(&H2713) Then
        SelectionCycle = "X"
    ElseIf CurrentValue = "X" Then
        SelectionCycle = ""
    End If
    
End Function

Public Function RowIsSyncSelected(Indicator As Variant) As Boolean
    If Indicator = ChrW(&H2713) Then RowIsSyncSelected = True
End Function

Public Sub SelectAllColumnsInRecord(SyncRows As Range, ColumnOffset As Long)
    
    'update: get the ColumnsOffset from the number of columns from sync/header row
    
    Application.EnableEvents = False
    SyncRows.Resize(SyncRows.Rows.Count, ColumnOffset).Select
    Application.EnableEvents = True
    
End Sub

Public Sub SelectorSelect(ws As Worksheet, Target As Range, rng)
    Dim IntersectRange  As Range
    Dim TargetText      As String
    Dim InteriorColor   As Long
    Dim FontColor       As Long
    
    Set IntersectRange = Application.Intersect(rng, Target)
    If IntersectRange Is Nothing Then
        Exit Sub
    End If
    
    With IntersectRange
        TargetText = .Value
        
        FontColor = vbBlack
        Select Case TargetText
            Case Is = "Update"  'orange
                InteriorColor = RGB(255, 229, 204)
            Case Is = "Trash"   'grey
                InteriorColor = RGB(180, 180, 180)
                FontColor = vbRed
            Case Is = "Restore" 'green
                InteriorColor = RGB(204, 255, 204)
            Case Is = "Delete"  'grey
                InteriorColor = RGB(180, 180, 180)
                FontColor = vbRed
            Case Else           'white
                InteriorColor = RGB(255, 255, 255)
        End Select
        
        .Font.Color = FontColor
        .Interior.Color = InteriorColor
        
    End With
    
End Sub

Public Sub SelectorSet()
    If Range("ViewTrash") Then
        Range("UpdateSelections").Cells(1) = "Restore"
        Range("UpdateSelections").Cells(2) = "Delete"
        Range("LogUpdateDelete") = "Restore"
    Else
        Range("UpdateSelections").Cells(1) = "Update"
        Range("UpdateSelections").Cells(2) = "Trash"
        Range("LogUpdateDelete") = "Update"
    End If
    Call XLFunc.SelectorSelect(Sheets("Log"), Range("LogUpdateDelete"), Range("LogUpdateDelete"))
End Sub

