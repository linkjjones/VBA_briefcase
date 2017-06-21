Attribute VB_Name = "XLFunc"
Option Explicit

'Libraries needed image
'BE SURE TO MAKE EVERY HEADER ROW THE SAME FOR EVERY PAGE!!!
Public Const HeaderRow As Long = 10
Public Const DataStartRow As Long = HeaderRow + 1
Public Const Orange = 46
Public pwd As String
Public Clean As Boolean
Public GlobalCounter As Long
Public Const SkipString As String = "~$"
Public Const NL As String = vbNewLine 'new line
Public Const DL As String = NL & NL   'skip a line

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

Public Function HasDependents(ByVal Target As Excel.Range) As Boolean
    On Error Resume Next
    HasDependents = Target.Dependents.Count
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

Public Function lastRow(ws As Worksheet, ColumnNumber As Long) As Long
    lastRow = ws.Cells(ws.Rows.Count, ColumnNumber).End(xlUp).Row
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

Public Sub CopyDownFormulas(ws As Worksheet, PasteRange As Range, FormulaCommentCell As Range, PasteAsValues As Boolean)
    
    ws.Unprotect
    
    If Not FormulaCommentCell.Comment Is Nothing Then
        'Make sure it starts with an '=' sign
        If Left(FormulaCommentCell.Comment.Text, 1) = "=" Then
            PasteRange = FormulaCommentCell.Comment.Text
            If PasteAsValues Then
                PasteRange.Value = PasteRange.Value
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
    XLFunc.CopyUpFormulas_Sheet
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
        EndColumn = XLFunc.GetLastCol(ws, DataStartRow)
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

Public Function GetLastCol(ws As Worksheet, Optional RowNumber As Long, _
                           Optional ColLimit As Long) As Long
    RowNumber = IIf(RowNumber = 0, HeaderRow, RowNumber)

    GetLastCol = ws.Cells(RowNumber, ws.Columns.Count).End(xlToLeft).Column
    
    GetLastCol = IIf(GetLastCol < ColLimit, ColLimit, GetLastCol)
End Function

Public Function GetFirstCol(ws As Worksheet, Optional RowNumber As Long, _
                            Optional StartCol As Long) As Long
    'This function assumes that the start cell is blank
    
    RowNumber = IIf(RowNumber = 0, HeaderRow, RowNumber)
    StartCol = IIf(StartCol = 0, 1, StartCol)
    
    GetFirstCol = ws.Cells(RowNumber, StartCol).End(xlToRight).Column
End Function

Public Function GetLastRow(ws As Worksheet, Optional ColumnNumber As Long, _
                           Optional LimitRow As Long, _
                           Optional ToColumn As Long) As Long
    Dim lastRow     As Long
    Dim i           As Long
    Dim ColLastRow  As Long
    
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
    
    GetLastRow = IIf(lastRow < LimitRow, LimitRow, lastRow)
    
End Function

'Public Function GetFirstRow(ws As Worksheet, ColumnNumber As Long, _
'                           Optional StartRow As Long, _
'                           Optional BtmLimitRow As Long) As Long
'    'Using XLDown is funky as it goes from where it is (a blank/not blank cell)
     'to the next DIFFERENT cell
'    StartRow = IIf(StartRow = 0, HeaderRow, StartRow)
'
'    GetFirstRow = ws.Cells(StartRow, ColumnNumber).End(xlDown).Row
'
'    If BtmLimitRow > 0 Then
'        GetFirstRow = IIf(GetFirstRow > BtmLimitRow, BtmLimitRow, GetFirstRow)
'    End If
'
'End Function

Public Function HeaderCol(ws As Worksheet, HeaderName As String, Optional HeadingRow As Long) _
                          As Long
    On Error Resume Next
    Dim Header As Range

    Dim LookRange As Range, cell As Range
    Dim LastDataColumn As Long
    
    HeadingRow = IIf(HeadingRow = 0, HeaderRow, HeadingRow)
    
    LastDataColumn = GetLastCol(ws, HeadingRow)
    
    With ws
        'since the above code is flaky...lets just loop through
        Set LookRange = .Range(.Cells(HeadingRow, 1), .Cells(HeadingRow, LastDataColumn))
        For Each cell In LookRange
            If cell.Value = HeaderName Then
                HeaderCol = cell.Column
                Exit For
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
    On Error Resume Next
    
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
    rng.Borders.ColorIndex = xlNone
End Sub

Public Sub FormatRangeWithLines(formatRange As Range, Optional VerticalLines As Boolean)
    
    With formatRange
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

Public Sub SelectionFormat(ws As Worksheet, Target As Range, rng As Range, _
                     UpdateText As String, _
                     Optional ValidationColOffset As Long)
    'This routine finds the intersection of
    Dim IntersectRange As Range

    Set IntersectRange = Application.Intersect(rng, Target)
    If IntersectRange Is Nothing Then
        Exit Sub
    End If
    
    With IntersectRange
        If IntersectRange.Cells(1).Value = "" Then
            If ValidationColOffset = 0 Or _
               IntersectRange.Cells(1).Offset(, ValidationColOffset).Value <> "" Then
                .Value = UpdateText
                If .Cells(1).Value = "Trash" Or .Cells(1).Value = "Delete" Then
                    .Interior.Color = vbBlack
                    .Font.Color = RGB(255, 0, 0)
'                    .Borders.Color = RGB(255, 0, 0)
                ElseIf .Cells(1).Value = "Update" Then
                    .Interior.ColorIndex = Orange
                    .Font.Color = RGB(255, 255, 255)
'                    .Borders.ColorIndex = xlNone
                ElseIf .Cells(1).Value = "Restore" Then
                    .Interior.Color = RGB(0, 153, 0)
                    .Font.Color = vbWhite
'                    .Borders.ColorIndex = xlNone
                End If
            End If
        Else
            .Value = ""
            .Interior.ColorIndex = 0
'            .Borders.ColorIndex = xlNone
        End If
    End With

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

Public Function GetAllFilesInFolder(HostFolderPath As String, _
                                    Optional ValidatorString As String, _
                                    Optional SkipFileString As String) As Collection
    Dim fso, oFolder, oSubfolder, oFile, queue As Collection
    Dim FileCollection As New Collection
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
    queue.add fso.GetFolder(HostFolderPath)

    Do While queue.Count > 0
        Set oFolder = queue(1)
        queue.Remove 1 'dequeue
        '...insert any folder processing code here...
        For Each oSubfolder In oFolder.SubFolders
            queue.add oSubfolder 'enqueue
        Next oSubfolder
        
        For Each oFile In oFolder.Files
            If InStr(1, oFile, ValidatorString) > 0 Then
                If SkipFileString <> "" Then
                    If InStr(1, oFile, SkipFileString) = 0 Then
                        FileCollection.add oFile
                    End If
                Else
                    FileCollection.add oFile
                End If
            End If
        Next oFile
        
    Loop
    
    Set GetAllFilesInFolder = FileCollection
    
End Function

Public Function CountB(rng As Range)
    Dim i As Long
'    Dim j As Long
    Dim c As Range
    
    For Each c In rng
        i = i + (c.Value <> "") * -1
'        j = j + 1
    Next c
    
'    Debug.Print j
    CountB = i
End Function

'Public Sub testInts()
'    Dim tmp As Collection: Set tmp = New Collection
'
'    tmp.add 3: tmp.add 1: tmp.add 4
'    'if next line (2) is commented out:     if dupes: "1,3,4,4"  if uniques: "1,3,4"
'    tmp.add 2                    'else:     if dupes: "1,2,3,4,4 if uniques: "1,2,3,4"
'    tmp.add 4
'    Set tmp = mergeSort(tmp, False)
'
'End Sub
'
'Public Sub testStrings()
'    Dim tmp As Collection: Set tmp = New Collection
'
'    tmp.add "C": tmp.add "A": tmp.add "D"
'    'if next line ("B") is commented out:   if dupes: "A,C,D,D"  if uniques: "A,C,D"
'    'tmp.Add "B"         'else:             if dupes: "A,B,C,D,D" if uniques: "A,B,C,D"
'    tmp.add "D"
'    Set tmp = mergeSort(tmp, False)
'End Sub

Public Function mergeSort(c As Collection, Optional uniq = True) As Collection

    Dim i As Long, xMax As Long, tmp1 As Collection, tmp2 As Collection, xOdd As Boolean

    Set tmp1 = New Collection
    Set tmp2 = New Collection

    If c.Count = 1 Then
        Set mergeSort = c
    Else

        xMax = c.Count
        xOdd = (c.Count Mod 2 = 0)
        xMax = (xMax / 2) + 0.1     ' 3 \ 2 = 1; 3 / 2 = 2; 0.1 to round up 2.5 to 3

        For i = 1 To xMax
            tmp1.add c.Item(i) & "" 'force numbers to string
            If (i < xMax) Or (i = xMax And xOdd) Then tmp2.add c.Item(i + xMax) & ""
        Next i

        Set tmp1 = mergeSort(tmp1, uniq)
        Set tmp2 = mergeSort(tmp2, uniq)

        Set mergeSort = merge(tmp1, tmp2, uniq)

    End If
    
End Function

Private Function merge(c1 As Collection, c2 As Collection, _
                       Optional ByVal uniq As Boolean = True) As Collection

    Dim tmp As Collection
    Set tmp = New Collection

    If uniq = True Then On Error Resume Next    'hide duplicate errors

    Do While c1.Count <> 0 And c2.Count <> 0
        If c1.Item(1) > c2.Item(1) Then
            If uniq Then tmp.add c2.Item(1), c2.Item(1) Else tmp.add c2.Item(1)
            c2.Remove 1
        Else
            If uniq Then tmp.add c1.Item(1), c1.Item(1) Else tmp.add c1.Item(1)
            c1.Remove 1
        End If
    Loop

    Do While c1.Count <> 0
        If uniq Then tmp.add c1.Item(1), c1.Item(1) Else tmp.add c1.Item(1)
        c1.Remove 1
    Loop
    Do While c2.Count <> 0
        If uniq Then tmp.add c2.Item(1), c2.Item(1) Else tmp.add c2.Item(1)
        c2.Remove 1
    Loop
    On Error GoTo 0

    Set merge = tmp

End Function

Public Sub HideAllColumns(ws As Worksheet, StartCol As Long, LastColumn As Long)
    Dim i As Long
    
    'Hide Columns
    For i = StartCol To LastColumn
        ws.Columns(i).EntireColumn.Hidden = True
    Next i
End Sub

Public Sub ProgressBar(Msg As String, Done As Long, Total As Long)
    'This function puts text into the statusbar
    Application.StatusBar = Msg & " [ " & String(Done, "|") & String(Total - Done, ".") & " ] " & Format(Done / Total, "Percent")
End Sub

Public Function LatestVersion() As Boolean
    If Not Range("LatestVersion") > Range("AboutVersion") Then
        LatestVersion = True
    End If
End Function
