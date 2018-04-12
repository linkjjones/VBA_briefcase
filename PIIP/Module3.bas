Attribute VB_Name = "Module3"
Public TML


Sub InsertMinTML()

If Left(Sheets("Homepage").Range("B5"), 2) = "SB" Then
    SBRowForm.Show
Else
    MinTPoint.Show
End If

End Sub

Sub SaveTemplateAs()

If ActiveSheet.Range("B1") = "" Or ActiveSheet.Range("B4") = "" Or ActiveSheet.Range("B5") = "" Then
    MsgBox "Please fill in the Inspection Date, Corrosion Group and Circuit before saving this workbook", vbCritical
    End
End If

templatename = Trim(ActiveSheet.Range("B4")) & " " & Trim(ActiveSheet.Range("B5")) & " " & ActiveSheet.Range("B1")

FName = Application.GetSaveAsFilename(InitialFileName:=templatename, filefilter:="Excel Files (*.xlsm), *.xlsm")
If FName = False Then End

useranswer = MsgBox("Would you like to sign this report on behalf of Acuren?", vbYesNo, "Acuren Signature")

If useranswer = vbYes Then
    For X = 1 To Sheets("Homepage").Shapes.Count
        If Sheets("Homepage").Shapes(X).Name = "AcurenSignature" Then
            MsgBox "You can only sign a template once.", vbCritical
            ActiveWorkbook.SaveAs Filename:=FName
            End
        End If
    Next X
    Sheets("Template").Shapes("AcurenSignature").Copy
    Sheets("Homepage").Range("E2").Select
    Sheets("Homepage").Paste
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = _
            "I, " & Application.UserName & " (" & GetLogonName & ") on " & Now() & ", certify that the information contained in this report is accurate to the best of my knowledge."
    Selection.Placement = xlFreeFloating
End If
ActiveWorkbook.SaveAs Filename:=FName



End Sub

Sub InsertTMLPhoto()

Dim pict As Variant
Dim ImgFileFormat As String
Dim dwgrange As Range
Dim comPic As Shape
Set dwgrange = Selection
Dim PicArr() As Variant
Dim PicCluster As Shape

ActiveSheet.Unprotect "Dh1986"

ImgFileFormat = "Image Files jpg (*.jpg), *.jpg, bmp (*.bmp), *.bmp, tif (*.tif),*.tif"
pict = Application.GetOpenFilename(ImgFileFormat, MultiSelect:=True)
On Error GoTo errHandler
If pict = False Then End 'this ends program when cancel button is pressed but gets an error when multiple images are selected

resumeInsert:

For X = LBound(pict) To UBound(pict)
    Set comPic = ActiveSheet.Shapes.AddPicture(pict(X), False, True, 0, 0, 100, 100)
    ReDim Preserve PicArr(X)
    PicArr(X) = ActiveSheet.Pictures(ActiveSheet.Pictures.Count).Name
    With comPic
        .OnAction = "ImageZoom"
        .LockAspectRatio = msoTrue
        .Height = dwgrange.Height
        .Width = dwgrange.Width - 20
            
        If .Height > dwgrange.Height Then .Height = dwgrange.Height - 10
            
        .Top = (dwgrange.Top + (dwgrange.Height / 2)) - (.Height / 2)
        If UBound(pict) = 1 Then
            .Left = dwgrange.Left + ((dwgrange.Width - .Width) / 2)
        Else
            If X = 1 Then
                .Left = dwgrange.Left + 10
            Else
                .Left = ActiveSheet.Pictures(ActiveSheet.Pictures.Count - 1).Left + ActiveSheet.Pictures(ActiveSheet.Pictures.Count - 1).Width + 10
            End If
        End If
    End With
    dwgrange.ClearContents
Next X
If UBound(PicArr) > 1 Then
    Set PicCluster = ActiveSheet.Shapes.Range(PicArr).Group
    PicCluster.Placement = xlFreeFloating
    If ActiveSheet.Columns(dwgrange.Column).Width < PicCluster.Width Then
        RequiredColumnWidthInPoints = PicCluster.Width
        PointsToCharactersRatio = ActiveSheet.Columns(1).Width / ActiveSheet.Columns(1).ColumnWidth
        ActiveSheet.Columns(dwgrange.Column).ColumnWidth = (RequiredColumnWidthInPoints / PointsToCharactersRatio) + 20
        
    End If
    
    With PicCluster
       .Left = dwgrange.Left + ((dwgrange.Width - .Width) / 2)
    End With
    PicCluster.Placement = xlMove
    PicCluster.OnAction = "ImageZoom"
End If

ActiveSheet.Protect "Dh1986"

errHandler:
If Err.Number = 13 Then
    Err.Clear
    GoTo resumeInsert:
End If

End Sub

Sub ImageZoom()

ActiveSheet.Shapes(Application.Caller).Select

If ActiveWindow.Zoom <= 100 Then
    ActiveWindow.Zoom = 180
Else
    ActiveWindow.Zoom = 100
End If
SendKeys "{ESC}"


'Range("A1").Select

End Sub

Sub ImageZoom2()

'You need to select a picture before running this code
'else it will give you error'
Dim TempChart As String, Picture2Export As String
Dim PicWidth As Long, PicHeight As Long
Dim FSO As Object
Set FSO = CreateObject("scripting.filesystemobject")

TempName = FSO.GetTempName
TempName = Left(TempName, Len(TempName) - 4) & ".gif"
ASheetName = ActiveSheet.Name
If ActiveWorkbook.Path <> "" Then FName = ActiveWorkbook.Path & TempName '& "\gif" '"\temp.gif"
If ActiveWorkbook.Path = "" Then FName = Application.DefaultFilePath & "\temp.gif"

ActiveSheet.Pictures(Application.Caller).Select

Picture2Export = Selection.Name
'Store the picture's height and width  in a variable
With Selection
    PicHeight = .ShapeRange.Height
    PicWidth = .ShapeRange.Width
End With
'Add a temporary chart in sheet1
Charts.Add
ActiveChart.Location Where:=xlLocationAsObject, Name:=ASheetName
'Selection.Border.LineStyle = 0
TempChart = Selection.Name & " " & Split(ActiveChart.Name, " ")(2)
With ActiveSheet
'Change the dimensions of the chart to suit your need
With .Shapes(TempChart)
.Width = PicWidth
.Height = PicHeight
End With
'Copy the picture
.Shapes(Picture2Export).Copy
'Paste the picture in the chart
With ActiveChart
'.ChartArea.Select
.Paste
End With
'Finally export the chart
.ChartObjects(1).Chart.Export Filename:=FName, FilterName:="GIF"
'Destroy the chart. You may want to delete it...
.Shapes(TempChart).Delete

PictureViewer.Image1.Picture = LoadPicture(FName)
Kill FName

End With

PictureViewer.Show
End Sub

Sub MakePlots()

Dim AlignedDataSet1() As Variant
Dim ChartName As String

If Sheets("Homepage").Range("A17") = "No Splits" Then
    For X = 1 To ActiveWorkbook.Sheets.Count
        If Sheets(X).Name <> "ListSheet" And Sheets(X).Name <> "Template" And Sheets(X).Name <> "BlankWS" And Sheets(X).Name <> "CalcSheet" And Sheets(X).Name <> "Homepage" And Sheets(X).Type <> 3 Then
            Y = 2
            Do
                If WorksheetFunction.IsNumber(Sheets(X).Cells(Y, 12)) Then
                    ReDim Preserve AlignedDataSet1(Z)
                    AlignedDataSet1(Z) = Sheets(X).Cells(Y, 12)
                    Z = Z + 1
                End If
                Y = Y + 1
            Loop Until Sheets(X).Cells(Y, 1) = ""
        End If
    Next X
    ChartName = Trim("Plot " & ActiveSheet.Range("A17"))
    If Z > 0 Then
        MakeExcelPPlot AlignedDataSet1, ChartName
    Else
        End
    End If
    Sheets("Homepage").Activate
Else
    For l = 0 To ActiveSheet.SplitOptionList.ListCount - 1
        If ActiveSheet.SplitOptionList.Selected(l) = True Then
            ReDim AlignedDataSet1(1)
            For X = 1 To ActiveWorkbook.Sheets.Count
                If Sheets(X).Name <> "ListSheet" And Sheets(X).Name <> "Template" And Sheets(X).Name <> "BlankWS" And Sheets(X).Name <> "CalcSheet" And Sheets(X).Name <> "Homepage" And Sheets(X).Type <> 3 Then
                    For c = 1 To 29
                        If Sheets(X).Cells(1, c) = ActiveSheet.Range("A17") Then Exit For
                    Next c
                    Y = 2
                    Do
                        If WorksheetFunction.IsNumber(Sheets(X).Cells(Y, 12)) And Left(Sheets(X).Cells(Y, c), Len(Sheets(X).Cells(Y, c))) = Left(ActiveSheet.SplitOptionList.List(l), Len(ActiveSheet.SplitOptionList.List(l))) Then
                            ReDim Preserve AlignedDataSet1(Z)
                            AlignedDataSet1(Z) = Sheets(X).Cells(Y, 12)
                            Z = Z + 1
                        End If
                        Y = Y + 1
                    Loop Until Sheets(X).Cells(Y, 1) = ""
                End If
            Next X
            ChartName = Trim("Plot " & ActiveSheet.Range("A17") & " " & ActiveSheet.SplitOptionList.List(l))
            If Z > 0 Then
                MakeExcelPPlot AlignedDataSet1, ChartName
            Else
                End
            End If
            Sheets("Homepage").Activate
        End If
    Next l
End If



End Sub

Sub PopulateListBox(SplitOption)
Dim splitArr() As Variant

'Set hp = Sheets("Homepage")

If SplitOption = "No Splits" Then
    ActiveSheet.SplitOptionList.Clear
    End
End If

s = 1
For X = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(X).Name <> "ListSheet" And Sheets(X).Name <> "Template" And Sheets(X).Name <> "BlankWS" And Sheets(X).Name <> "CalcSheet" And Sheets(X).Name <> "Homepage" And Sheets(X).Type <> 3 Then
        For Y = 1 To 29
            If Sheets(X).Cells(1, Y) = SplitOption Then Exit For
        Next Y
        r = 2
        Do
            ReDim Preserve splitArr(s)
            splitArr(s) = Sheets(X).Cells(r, Y)
            r = r + 1
            s = s + 1
        Loop Until Sheets(X).Cells(r, 1) = ""
    End If
Next X

If r > 0 Then
    removeDuplicates splitArr
Else
    End
End If

For X = 1 To UBound(splitArr)
    If splitArr(X) <> "" Then ActiveSheet.SplitOptionList.AddItem splitArr(X)
Next X

End Sub
