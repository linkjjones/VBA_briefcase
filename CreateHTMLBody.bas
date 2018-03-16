Attribute VB_Name = "CreateHTMLBody"
Option Explicit

Function RangetoHTML(rng As Range)
' By Ron de Bruin.
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
'    Dim TempWB As Workbook
    Dim WB As Workbook
    
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

'    'Copy the range and create a new workbook to past the data in
'    rng.Copy
'    Set TempWB = Workbooks.add(1)
'    With TempWB.Sheets(1)
'        .Cells(1).PasteSpecial Paste:=8
'        .Cells(1).PasteSpecial xlPasteValues, , False, False
'        .Cells(1).PasteSpecial xlPasteFormats, , False, False
'        .Cells(1).Select
'        Application.CutCopyMode = False
'        On Error Resume Next
'        .DrawingObjects.Visible = True
'        .DrawingObjects.Delete
'        On Error GoTo 0
'    End With

'    'Publish the sheet to a htm file
'    With TempWB.PublishObjects.add( _
'         SourceType:=xlSourceRange, _
'         Filename:=TempFile, _
'         Sheet:=TempWB.Sheets(1).Name, _
'         Source:=TempWB.Sheets(1).UsedRange.Address, _
'         HtmlType:=xlHtmlStatic)
'        .Publish (True)
'    End With
    
    Set WB = ActiveWorkbook
    'Publish the sheet to a htm file
    With WB.PublishObjects.add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:="Email Format", _
         Source:=Range("EmailFormat").Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    
    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

'    'Close TempWB
'    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
'    Set TempWB = Nothing
    Set WB = Nothing
    
End Function

