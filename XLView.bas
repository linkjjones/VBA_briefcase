Attribute VB_Name = "XLView"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Public Sub FreezePanes(ws As Worksheet, _
                       Optional SplitRow As Long, _
                       Optional SplitColumn As Long)
    Dim CurrentSheet As Worksheet
    
    If SplitRow = 0 Then
        SplitRow = DataStartRow
    End If
    
     '// store current sheet
    Set CurrentSheet = ActiveSheet
     
     '// Stop flickering...
    Application.ScreenUpdating = False
              
     '// Have to activate - SplitColumn and SplitRow are properties
     '// of ActiveSheet
    ws.Activate
     
    With ActiveWindow
        'reset: may not be necessary
        .FreezePanes = False
        .SplitColumn = SplitColumn
        .SplitRow = SplitRow
        .FreezePanes = True
    End With
     
     '// Back to original sheet
    CurrentSheet.Activate
    Application.ScreenUpdating = True
     
    Set ws = Nothing
    Set CurrentSheet = Nothing
     
End Sub

Public Sub UnfilterSheet(Optional ws As Worksheet)
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If ws.FilterMode Then ws.ShowAllData
End Sub

'Public Sub ScrollToCol(ScrollCol As Integer)
'    ActiveWindow.ScrollColumn = ScrollCol
'End Sub

Public Sub ScrollToCol(ScrollCol As Integer, Optional SmoothUP As Boolean)
    Dim i As Integer
    Dim StartingCol As Long
    Dim startTime As Long
    On Error GoTo NormalScroll
    
    Application.EnableEvents = False
    Sleep 1
    
    With ActiveWindow
        StartingCol = .ScrollRow
        If SmoothUP Then
            For i = StartingCol To ScrollCol
                .ScrollColumn = i
                Sleep Inertia(Round((i - StartingCol + 1) / (ScrollCol - StartingCol), 2))
            Next i
        Else
           GoTo NormalScroll
        End If
    End With
    
    Application.EnableEvents = True
    
Exit Sub

NormalScroll:
    If ScrollCol > 0 Then
        ActiveWindow.ScrollColumn = ScrollCol
    End If
    Application.EnableEvents = True
    
End Sub

Public Sub ScrollToRow(ScrollRow As Integer, Optional SmoothUP As Boolean)
    Dim i As Integer
    Dim StartingRow As Long
    Dim startTime As Long
    On Error GoTo NormalScroll
    
    Application.EnableEvents = False
    
    Sleep 1
    
    With ActiveWindow
        StartingRow = .ScrollRow
        If SmoothUP Then
            For i = StartingRow To ScrollRow
                .ScrollRow = i
                Sleep Inertia(Round((i - StartingRow + 1) / (ScrollRow - StartingRow), 2))
            Next i
        Else
           GoTo NormalScroll
        End If
    End With
    
    Application.EnableEvents = True
    
Exit Sub

NormalScroll:
    If ScrollRow > 0 Then
        ActiveWindow.ScrollRow = ScrollRow
    End If
    Application.EnableEvents = True
    
End Sub

Private Function Inertia(PercentScrolled As Double) As Integer
    Dim k As Integer
    Dim h As Integer
    Dim a As Integer
    
    a = 200 'rate of change
    h = -0.1 'negative y: constant (0 to 1: 50%=0.5)
    k = 3  'x fastest (midpoint)
    
    Inertia = Round(a * (PercentScrolled + h) ^ 2 + k)
'    Debug.Print Inertia
End Function

Public Sub HideAllXLControls()
    Dim ws As Worksheet
    Dim CurrentSheet As Worksheet
    
    'Get the current ws so we can go back to it after all the changes
    Set CurrentSheet = ActiveSheet
    
    With Application
        .DisplayFormulaBar = False
'        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
'        .DisplayScrollBars = False
'        .DisplayStatusBar = Not Application.DisplayStatusBar
    End With
    
    'This is only collapses the ribbon
    If CommandBars("Ribbon").Height > 100 Then
        CommandBars.ExecuteMso "MinimizeRibbon"
    End If
    
    Application.ScreenUpdating = False
    For Each ws In Worksheets
        ws.Select
        With ActiveWindow
            .DisplayWorkbookTabs = False
            .DisplayHeadings = False
            .DisplayGridlines = False
'            .Zoom = Worksheets("Control").Range("WBZoom").Value
        End With
    Next ws
    
    'Go back to the starting worksheet
    CurrentSheet.Select
    
    Application.ScreenUpdating = True
   
End Sub

Public Sub ShowAllXLControls()
    Dim ws As Worksheet
    Dim CurrentSheet As Worksheet
    
    'Get the current ws so we can go back to it after all the changes
    Set CurrentSheet = ActiveSheet
    
    Application.ScreenUpdating = False
    
    'Set zoom for each worksheet
    For Each ws In Worksheets
        ws.Select
        With ActiveWindow
'            .DisplayGridlines = True
            .DisplayHeadings = True
            .DisplayWorkbookTabs = True
        End With
    Next ws
    
    'maximize ribbon
    If CommandBars("Ribbon").Height < 100 Then
        CommandBars.ExecuteMso "MinimizeRibbon"
    End If
    
    With Application
        .DisplayFormulaBar = True
'        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        .DisplayScrollBars = True
        .DisplayStatusBar = True
    End With
    
    'Go back to the starting worksheet
    CurrentSheet.Select
    
    Application.ScreenUpdating = True
    
 End Sub

Public Sub HideAllColumns(ws As Worksheet, StartCol As Long, LastColumn As Long)
    Dim i As Long
    
    'Hide Columns
    For i = StartCol To LastColumn
        ws.Columns(i).EntireColumn.Hidden = True
    Next i
End Sub

Public Sub ProgressBar(Msg As String, Optional Done As Long, Optional Total As Long)
    Dim percentDone As Double
    
'    On Error Resume Next
    If Msg <> "" Then
        If Done <> 0 And Total <> 0 Then
            percentDone = Round(Done / Total * 100)
            Msg = Msg & " [ " & String(percentDone, "|") & String(100 - percentDone, ".") & " ] " & Format(Done / Total, "Percent")
        End If
        If Msg <> Application.StatusBar Then
            Application.StatusBar = Msg
        End If
    Else
        Application.StatusBar = False
    End If
    
End Sub

'Public Sub ProgressBar(Msg As String, Done As Long, Total As Long)
'    On Error Resume Next
'    'USE PERCENTAGE INSTEAD OF ACTUAL PASSED NUMBERS
'    Application.StatusBar = Msg & " [ " & String(Done, "|") & String(Total - Done, ".") & " ] " '& Format(Done / Total, "Percent")
'End Sub
