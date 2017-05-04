Attribute VB_Name = "XLFunctions"
Option Explicit

'Libraries needed image
'BE SURE TO MAKE EVERY HEADER ROW THE SAME FOR EVERY PAGE!!!

Public Const HeaderRow As Long = 10
Public Const DataStartRow As Long = HeaderRow + 1
Public Const Orange = 46
Public pwd As String
Public Clean As Boolean

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub SetPassword()
    pwd = ""
End Sub

Public Sub InsertDateNow(DateCell As Range)
    DateCell.Value = Date
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
    Dim startTime As Long
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
    Dim XLApp As Excel.Application
    Dim XLBook As Excel.Workbook
    
    'Open Spreadsheet
    Set XLApp = CreateObject("Excel.application")

OpenXLBook:
    XLApp.Application.AskToUpdateLinks = False
    XLApp.Application.DisplayAlerts = False
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
    XLApp.Application.AskToUpdateLinks = True
    XLApp.Application.DisplayAlerts = True
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
    Dim XLApp As Excel.Application
    
    Set XLApp = XLBook.Application
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

Public Sub CopyDownFormulas(ws As Worksheet, pasteRange As Range, FormulaCommentCell As Range, PasteAsValues As Boolean)
    
    ws.Unprotect
    
    If Not FormulaCommentCell.Comment Is Nothing Then
        'Make sure it starts with an '=' sign
        If Left(FormulaCommentCell.Comment.Text, 1) = "=" Then
            pasteRange = FormulaCommentCell.Comment.Text
            If PasteAsValues Then
                pasteRange.Value = pasteRange.Value
            End If
        End If
    End If
 
End Sub

Public Sub CopyDownFormulas_Sheet(LastDataRow As Long, _
                                  Optional ws As Worksheet, _
                                  Optional CommentRow As Long, _
                                  Optional FirstDataRow As Long, _
                                  Optional LeaveAsFormulas As Boolean)
    Dim i As Long
    Dim CopyRange As Range
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    If FirstDataRow = 0 Then
        FirstDataRow = DataStartRow
    End If
    If CommentRow = 0 Then
        CommentRow = HeaderRow
    End If
    
    ws.Unprotect
    
    With ws
        For i = 1 To GetLastCol(ws, CommentRow)
            If Not .Cells(CommentRow, i).Comment Is Nothing Then
                'Make sure it starts with an '=' sign
                If Left(.Cells(CommentRow, i).Comment.Text, 1) = "=" Then
                    Set CopyRange = .Range(.Cells(FirstDataRow, i), .Cells(LastDataRow, i))
                    CopyRange = .Cells(CommentRow, i).Comment.Text
                    If Not LeaveAsFormulas Then
                        CopyRange.Value = CopyRange.Value
                    End If
                End If
            End If
        Next i
    End With

End Sub

Sub TestCopyUpFormulas()
    XLFunctions.CopyUpFormulas_Sheet
End Sub

Sub CopyUpFormulas_Sheet(Optional ws As Worksheet, Optional CommentRow As Long, _
                         Optional StartColumn As Long, Optional EndColumn As Long)
    'Loop through collumns and copy up to row above into comments
    'ASSUMES THE DataStartRow IS SET
    Dim i As Long
    Dim c As Range
    On Error Resume Next
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    If CommentRow = 0 Then
        CommentRow = DataStartRow - 1
    End If
    
    If StartColumn = 0 Then
        StartColumn = 1
    End If
    
    If EndColumn = 0 Then
        EndColumn = XLFunctions.GetLastCol(ws, DataStartRow)
    End If
    
    For i = StartColumn To EndColumn
        Set c = ws.Cells(DataStartRow, i)
        c.Select
        If Left(c.Formula, 1) = "=" Then
        'this is a formula, so copy to comments
            ws.Cells(CommentRow, i).Comment.Delete
            ws.Cells(CommentRow, i).AddComment CStr(c.Formula)
        End If
    Next i
    
End Sub

Public Sub ClearListBoxSelection(lst As msforms.ListBox)
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
    Dim LastRow As Long
    
    LastRow = ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).Row
    
    If LimitRow <> 0 Then
        LastRow = IIf(LastRow < LimitRow, LimitRow, LastRow)
    End If
    
    GetLastRow = LastRow
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
    
    With Application
        .DisplayFormulaBar = False
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
'        .DisplayScrollBars = False
'        .DisplayStatusBar = Not Application.DisplayStatusBar
    End With
    
    Application.ScreenUpdating = False
    'Set zoom for each worksheet
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
    currentSheet.Select
    
    Application.ScreenUpdating = True
   
End Sub

Public Sub ShowAllXLControls()
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    
    'Get the current ws so we can go back to it after all the changes
    Set currentSheet = ActiveSheet
    
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
    
    With Application
        .DisplayFormulaBar = True
        .ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        .DisplayScrollBars = True
        .DisplayStatusBar = True
    End With
    
    'Go back to the starting worksheet
    currentSheet.Select
    
    Application.ScreenUpdating = True
    
 End Sub
 
'Public Sub testthisInsert()
'   Call AddCheckBoxesToRange(Sheets("Control"), Range("B11:B25"))
'End Sub
'Public Sub testthisRemove()
'   Call RemoveCheckBoxesFromRange(Sheets("Control"), Range("B11:B25"))
'End Sub
 
Public Sub AddCheckBoxesToRange(ws As Worksheet, CheckRange As Range, _
                                Optional SelectCellRowOffset As Long, _
                                Optional SelectCellColumnOffset As Long, _
                                Optional ImageToRotate As String, _
                                Optional Increment As Integer)
    Dim c As Range
'    On Error Resume Next
    
    For Each c In CheckRange.Cells
        'rotate an image
        If ImageToRotate <> "" Then
            RotateImage ws, ImageToRotate, Increment, c.Row
        End If
        
        ws.CheckBoxes.add(c.Left, c.Top, c.Width, c.Height).Select
        With Selection
            .LinkedCell = c.Address
            .Characters.Text = ""
            .Name = c.Address
            c.Value = False
            c.Font.Color = vbWhite
        End With
    Next c
    
    'Deselect the last checkbox so it doesn't screw up other routines
    CheckRange.Cells(1 + SelectCellRowOffset, 1 + SelectCellColumnOffset).Select
    
 End Sub
 
 Public Sub RemoveCheckBoxesFromRange(ws As Worksheet, CheckRange As Range)
    On Error Resume Next
    Dim c As Object
    
    For Each c In ws.CheckBoxes
        If Not Intersect(c.TopLeftCell, CheckRange) Is Nothing Then
            c.Delete
        End If
    Next

 End Sub

'Sub testthis()
'    Call RemoveCheckBoxesFromSheet(Sheets("Control"))
'End Sub

 Public Sub RemoveCheckBoxesFromSheet(ws As Worksheet)
    On Error Resume Next
    Dim c As Object
    
    For Each c In ws.CheckBoxes
        c.Delete
    Next

 End Sub

Public Sub RotateImage(ws As Worksheet, imgName As String, _
                       Increment As Integer, Counter As Long)
    Dim Theta As Double
    Dim Position As Integer
    
    Position = Counter Mod Increment
    
    Theta = 360 / Increment
    DoEvents
    With ws.Shapes(imgName)
        .Rotation = Theta * Position
    End With
    
End Sub

Public Sub ValidateRange(RangeToBeValidated As Range, ValueListFormula As Variant, Optional DropDown As Boolean)
    With RangeToBeValidated.Validation
        .Delete
        .add xlValidateList, xlValidAlertStop, , ValueListFormula
        .IgnoreBlank = True
        .InCellDropdown = DropDown
    End With
End Sub

Public Sub ValidateRangeTrueFalse(RangeToBeValidated As Range, Optional DropDown As Boolean = True)
    With RangeToBeValidated.Validation
        .Delete
        .add xlValidateList, xlValidAlertStop, , "True, False"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Public Sub ClearLines(rng As Range)
    rng.Borders.LineStyle = xlNone
End Sub

Public Sub FormatRangeWithLines(FormatRange As Range, Optional VerticalLines As Boolean)
    
    With FormatRange
        If VerticalLines Then
            .Borders(xlInsideHorizontal).Color = RGB(200, 200, 200)
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlInsideVertical).Color = RGB(200, 200, 200)
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(100, 100, 100)
        Else
            .Borders(xlInsideHorizontal).Color = RGB(200, 200, 200)
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlEdgeBottom).Color = RGB(100, 100, 100)
        End If
    End With
    
End Sub

Public Function InRange(Range1 As Range, Range2 As Range) As Boolean
    ' returns True if Range1 is within Range2
    InRange = Not (Application.Intersect(Range1, Range2) Is Nothing)
End Function

Public Sub Selection(ws As Worksheet, target As Range, rng As Range)
    'This routine finds the intersection of
    Dim IntersectRange As Range

    Set IntersectRange = Application.Intersect(rng, target)
    If IntersectRange Is Nothing Then
        Exit Sub
    End If
    
    With IntersectRange
        If IntersectRange.Cells(1).Value = "" Then
            .Value = "Update"
            .Interior.ColorIndex = Orange
            .Font.Color = RGB(255, 255, 255)
        Else
            .Value = ""
            .Interior.ColorIndex = 0
        End If
    End With

End Sub

Function StringContainsNumber(strData As String) As Boolean
    Dim i As Integer
     
    For i = 1 To Len(strData)
        If IsNumeric(Mid(strData, i, 1)) Then
            StringContainsNumber = True
            Exit Function
        End If
    Next i
     
End Function


Sub FormatDecimal(rng As Range)
With rng
        .NumberFormat = "0.000"
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
    End With
End Sub

Sub FormatPercent(rng As Range)
With rng
        .NumberFormat = "0%"
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub FormatText(rng As Range)
With rng
        .NumberFormat = "@"
        .HorizontalAlignment = xlLeft
    End With
End Sub

Sub FormatID(rng As Range)
    With rng
        .Font.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
        .Font.Size = 8
    End With
End Sub

