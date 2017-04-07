Attribute VB_Name = "XLFunctions"
Option Explicit

'Libraries needed image
'BE SURE TO MAKE EVERY HEADER ROW THE SAME FOR EVERY PAGE!!!

Public Const HeaderRow As Long = 2
Public pwd As String
Public Clean As Boolean

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub SetPassword()
    pwd = ""
End Sub

Public Sub InsertDateNow(DateCell As Range)
    DateCell.Value = Date
End Sub

Public Sub UnfilterSheet()
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
End Sub

'Public Sub ScrollToCol(ScrollCol As Integer)
'    ActiveWindow.ScrollColumn = ScrollCol
'End Sub

Public Sub ScrollToCol(ScrollCol As Integer, Optional SmoothUP As Boolean)
    Dim i As Integer
    Dim StartingCol As Long
    Dim StartTime As Long
    On Error GoTo NormalScroll
    
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

Exit Sub

NormalScroll:
    If ScrollCol > 0 Then
        ActiveWindow.ScrollColumn = ScrollCol
   End If
    
End Sub

Public Sub ScrollToRow(ScrollRow As Integer, Optional SmoothUP As Boolean)
    Dim i As Integer
    Dim StartingRow As Long
    Dim StartTime As Long
    On Error GoTo NormalScroll
    
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

Exit Sub

NormalScroll:
    If ScrollRow > 0 Then
        ActiveWindow.ScrollRow = ScrollRow
   End If
    
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

Public Function HasDependents(ByVal target As Excel.Range) As Boolean
    On Error Resume Next
    HasDependents = target.Dependents.Count
End Function

Public Function OpenWorkbook(FilePath As String, Visible As Boolean, _
                             Optional Password As String, _
                             Optional WriteMode As Boolean) As Excel.Workbook
    On Error GoTo errHandler
    Dim XLApp As Excel.APPLICATION
    Dim XLBook As Excel.Workbook
    
    'Open Spreadsheet
    Set XLApp = CreateObject("Excel.application")

OpenXLBook:
    XLApp.APPLICATION.AskToUpdateLinks = False
    XLApp.APPLICATION.DisplayAlerts = False
    If WriteMode Then
        If Password = "" Or IsNull(Password) Then
            Set XLBook = XLApp.Workbooks.Open(FilePath, , False, , , , True, , , True)
        Else
            Set XLBook = XLApp.Workbooks.Open(FilePath, , False, , Password, , True, , , True)
        End If
    Else
        If Password = "" Or IsNull(Password) Then
            Set XLBook = XLApp.Workbooks.Open(FilePath, , True, , , , True)
        Else
            Set XLBook = XLApp.Workbooks.Open(FilePath, , True, , Password, , True)
        End If
    End If
    XLApp.APPLICATION.AskToUpdateLinks = True
    XLApp.APPLICATION.DisplayAlerts = True
    XLApp.Visible = Visible
    
    Set OpenWorkbook = XLBook
    
Exit Function
errHandler:
If Err.Number = 1004 Then
    MsgBox "Cannot access file: " & Chr(10) & FilePath, vbInformation
    XLApp.Quit
End If

End Function

Public Sub CloseWorkbook(XLBook As Excel.Workbook, Optional SaveWB As Boolean)
    Dim XLApp As Excel.APPLICATION
    
    Set XLApp = XLBook.APPLICATION
    If Not XLBook Is Nothing Then
        If SaveWB Then
            XLBook.Save
        Else
            XLBook.Saved = True
        End If
    End If
    XLApp.Quit
    Set XLApp = Nothing
    
End Sub

Public Function LastRow(ws As Worksheet, ColumnNumber As Long) As Long
    LastRow = ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).Row
End Function

Public Function LastCol(ws As Worksheet, RowNumber As Long) As Long
    LastCol = ws.Cells(RowNumber, ws.Columns.Count).End(xlToLeft).Column
End Function

Public Sub FormatRange(ws As Worksheet, StartCell As Range, EndCell As Range)
    'formating
    Dim i As Long
    Dim DataRow As Range
    
    'Clear all background, upper/lower line
    
    For i = StartCell.Row To EndCell.Row
        'Format entire row: Nav/Component
        'set range (row)
        Set DataRow = Range(ws.Cells(i, StartCell.Column), ws.Cells(i, EndCell.Column))
        If i Mod 2 > 0 Then
            'background
            DataRow.Interior.Color = RGB(180, 180, 180)
            'upper/lower line
            
        End If
        'Format vertical lines
        'Format Component
        'Format Material
        'Format Diameter
        
    Next i


End Sub

Public Function HeaderCell(HeaderName As String, ws As Worksheet, LastDataColumn As Long) As Range
    On Error Resume Next
    Dim Header As Range
    Dim LookRange As Range, cell As Range
    
    With ws
        'Get range last row in headers
'        Set HeaderCell = .Range(.Cells(HeaderRow, 1), .Cells(HeaderRow, LastDataColumn)).Find(HeaderName, , , xlWhole, , , True)
        'since the above code is flaky...lets just loop through
        Set LookRange = .Range(.Cells(HeaderRow, 1), .Cells(HeaderRow, LastDataColumn))
        For Each cell In LookRange
            If cell.Value = HeaderName Then
                Set HeaderCell = cell
            End If
        Next cell
        
    End With

End Function

Public Function WorksheetExists(wb As Workbook, SheetName As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
        If UCase(ws.Name) = UCase(SheetName) Then
            WorksheetExists = True
            Exit For
        End If
    Next

End Function

Public Function FindRowInColumn(XLSheet As Worksheet, FindCol As Long, _
                RowStart As Long, RowEnd As Long, FindValue As Variant) As Long
'---> Add reverse loop option
    Dim i As Long
    
    FindRowInColumn = 0
    With XLSheet
        For i = RowStart To RowEnd
            If FindValue = .Cells(i, FindCol) Then
'                Debug.Print .Cells(i, FindCol)
                FindRowInColumn = i
                Exit For
            End If
        Next i
    End With
    
End Function

Public Function FindColInRow(XLSheet As Worksheet, FindRow As Long, _
            ColStart As Long, ColEnd As Long, FindValue As Variant) As Long
'---> Add reverse loop option
    Dim i As Long
    
    FindColInRow = 0
    With XLSheet
        For i = ColStart To ColEnd
            If FindValue = .Cells(FindRow, i) Then
                FindColInRow = i
                Exit For
            End If
        Next i
    End With
    
End Function

Public Sub CopyDownFormulas(ws As Worksheet, DataStartRow As Long, LastDataRow As Long, PasteAsValue As Boolean)
    Dim i As Long
    Dim CopyRange As Range
    
    ws.Unprotect
    
    With ws
        For i = 1 To GetLastCol(ws, HeaderRow)
            If Not .Cells(HeaderRow, i).Comment Is Nothing Then
                If Left(.Cells(HeaderRow, i).Comment.Text, 1) = "=" Then
                    Set CopyRange = .Range(.Cells(DataStartRow, i), .Cells(LastDataRow, i))
                    CopyRange = .Cells(HeaderRow, i).Comment.Text
                    If PasteAsValue Then
                        CopyRange.Value = CopyRange.Value
                    End If
                End If
            End If
        Next i
    End With
    
    ws.Protect
    
End Sub

Public Sub ClearListBoxSelection(lst As MSForms.ListBox)
    Dim i As Integer
    
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = False
    Next i
    'Go back to top of list
    lst.Selected(0) = False
End Sub

Public Function GetLastCol(ws As Worksheet, RowNumber As Long) As Long
    If RowNumber = 0 Then
        GetLastCol = 1
        Exit Function
    End If
    GetLastCol = ws.Cells(RowNumber, ws.Columns.Count).End(xlToLeft).Column
End Function

Public Function GetLastRow(ws As Worksheet, ColumnNumber As Long, _
                           Optional LimitRow As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).Row
    
    If LimitRow <> 0 Then
        GetLastRow = IIf(GetLastRow < LimitRow, LimitRow, GetLastRow)
    End If
End Function

Public Function HeaderCol(ws As Worksheet, HeaderName As String) As Long
    On Error Resume Next
    Dim Header As Range

    Dim LookRange As Range, cell As Range
    Dim LastDataColumn As Long
    
    LastDataColumn = GetLastCol(ws, HeaderRow)
    
    With ws
        'since the above code is flaky...lets just loop through
        Set LookRange = .Range(.Cells(HeaderRow, 1), .Cells(HeaderRow, LastDataColumn))
        For Each cell In LookRange
            If cell.Value = HeaderName Then
                HeaderCol = cell.Column
                Exit Function
            End If
        Next cell
        
    End With

End Function

Public Function PDFExport(PrintRange As Range, FilePath_Name_ext As String) As Boolean
    PrintRange.ExportAsFixedFormat xlTypePDF, FilePath_Name_ext, xlQualityStandard, _
                                   True, False, , , True
End Function

Public Function IsValidFileName(NameToCheck As String) As Boolean
    If NameToCheck = "" Then Exit Function
    If InStr(1, NameToCheck, ":") > 0 Then Exit Function
    If InStr(1, NameToCheck, "?") > 0 Then Exit Function
    If InStr(1, NameToCheck, "<") > 0 Then Exit Function
    If InStr(1, NameToCheck, ">") > 0 Then Exit Function
    If InStr(1, NameToCheck, "/") > 0 Then Exit Function
    If InStr(1, NameToCheck, "\") > 0 Then Exit Function
    
    IsValidFileName = True
    
End Function

Public Sub SayThis(Sentence As String)
    Dim s As Object
    Dim vol As Long
    
    Set s = CreateObject("SAPI.SpVoice")
    
    'Get current volume
    vol = s.volume
    'set higher volume
    s.volume = 100
    'Say it
    s.Speak Sentence
    'Set volume to original level
    s.volume = vol
    'Cleanup
    Set s = Nothing
    
End Sub

Public Sub ExitEditMode()
'    If Application.EditDirectlyInCell = True Then
'        Application.EditDirectlyInCell = False
'    End If
End Sub

Public Sub CleanWorkbook()
    If Clean Then
        Call ShowAllXLControls
        Clean = False
    Else
        Call HideAllXLControls
        Clean = True
    End If
End Sub

Public Sub HideAllXLControls()
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    
    'Get the current ws so we can go back to it after all the changes
    Set currentSheet = ActiveSheet
    
    With APPLICATION
        .DisplayFormulaBar = False
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
        .DisplayScrollBars = False
        .DisplayStatusBar = Not APPLICATION.DisplayStatusBar
    End With
    
    APPLICATION.ScreenUpdating = False
    'Set zoom for each worksheet
    For Each ws In Worksheets
        ws.Select
        With ActiveWindow
            .DisplayWorkbookTabs = False
            .DisplayHeadings = False
            .DisplayGridlines = False
            .Zoom = Worksheets("Control").Range("WBZoom").Value
        End With
    Next ws
    
    'Go back to the starting worksheet
    currentSheet.Select
    
    APPLICATION.ScreenUpdating = True
   
End Sub

Public Sub ShowAllXLControls()
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    
    'Get the current ws so we can go back to it after all the changes
    Set currentSheet = ActiveSheet
    
    APPLICATION.ScreenUpdating = False
    
    'Set zoom for each worksheet
    For Each ws In Worksheets
        ws.Select
        With ActiveWindow
            .DisplayGridlines = True
            .DisplayHeadings = True
            .DisplayWorkbookTabs = True
        End With
    Next ws
    
    With APPLICATION
        .DisplayFormulaBar = True
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        .DisplayScrollBars = True
        .DisplayStatusBar = True
    End With
    
    'Go back to the starting worksheet
    currentSheet.Select
    
    APPLICATION.ScreenUpdating = True
    
 End Sub
