Attribute VB_Name = "Module7"
Option Base 1
Sub BandCharts(Target As Range)

Dim aSh As Worksheet
Dim PointArr() As Variant
Dim ThickArr() As Variant
Dim OrigDateArr() As Date
Dim LastInspectDateArr() As Variant
Dim OrigThickArr() As Variant
Dim LastInspectThick() As Variant
Dim InspectDateArr() As Date
Dim XPointArr(80) As Variant
Dim XThickArr(80) As Variant
Dim BandChart As Chart
Dim MaxThick As Double
Dim MaxThickArr(80) As Variant
Dim MaxReading As Double
Dim MinReading As Double
Dim FSO As Object



'trigger when A2 Check cell double clicked
'create band charts for each band of TML

'******Get data into array
'Loop 1: beginning cell to end cell of TML
'Loop 2: beginning cell to end cell of each band within TML

'******Expand data into rounder bands
'10 intervals between points (smoothing)

'******Plot data on spider charts
'Dataset1: OD (dark area)
'Dataset2: ID (white area)
'Title: Band letter + Weight Guess
'Save as jpg
'Load into form (all bands for the TML)

Set aSh = ActiveSheet
Set FSO = CreateObject("scripting.filesystemobject")

aSh.Unprotect "Dh1986"

r = Target.Row
'loop to find first row of TML clicked
If Target.Row > 2 Then
    Do Until Left(aSh.Cells(Target.Row, 1), Len(aSh.Cells(Target.Row, 1)) - 2) <> Left(aSh.Cells(r, 1), Len(aSh.Cells(r, 1)) - 2)
        r = r - 1
    Loop
End If

If r = 1 Then r = 2
If r > 2 Then r = r + 1


Do
        bc = bc + 1
        Erase PointArr
        Erase ThickArr
        Erase XPointArr
        Erase XThickArr
        Erase MaxThickArr
        Erase OrigThickArr
        Erase OrigDateArr
        Erase LastInspectThick
        Erase LastInspectDateArr
        Erase InspectDateArr
        
        MaxThick = 0
        MinReading = 10
        MaxReading = 0
        X = 0
        s = r
        
        'loop until last row of Band
        Do
                X = X + 1
                ReDim Preserve PointArr(X)
                ReDim Preserve ThickArr(X)
                ReDim Preserve OrigThickArr(X)
                ReDim Preserve OrigDateArr(X)
                ReDim Preserve LastInspectThick(X)
                ReDim Preserve LastInspectDateArr(X)
                ReDim Preserve InspectDateArr(X)
                
                If aSh.Cells(r, 12) <> "" Then
                    PointArr(X) = Right(aSh.Cells(r, 1), 2)
                    ThickArr(X) = aSh.Cells(r, 12)
                    OrigThickArr(X) = aSh.Cells(r, 5)
                    OrigDateArr(X) = aSh.Cells(r, 4)
                    LastInspectThick(X) = aSh.Cells(r, 7)
                    LastInspectDateArr(X) = aSh.Cells(r, 6)
                    InspectDateArr(X) = aSh.Cells(r, 8)
                    
                    If ThickArr(X) > MaxReading Then MaxReading = ThickArr(X)
                    If ThickArr(X) < MinReading Then MinReading = ThickArr(X)
                    
                Else
                    ThickArr(X) = 0
                    OrigThickArr(X) = 0
                    OrigDateArr(X) = 0
                    LastInspectThick(X) = 0
                    LastInspectDateArr(X) = 0
                    InspectDateArr(X) = 0
                End If
            r = r + 1
            If aSh.Cells(r, 1) = "" Then Exit Do
        'Loop Until Band Letter changes OR TML Changes
        Loop Until Left(Right(aSh.Cells(r, 1), 2), 1) <> Left(Right(aSh.Cells(s, 1), 2), 1) Or Left(aSh.Cells(r, 1), Len(aSh.Cells(r, 1)) - 2) <> Left(aSh.Cells(s, 1), Len(aSh.Cells(s, 1)) - 2) 'until next band/tml combination is detected
        
        'enter CRs into form list boxes
        On Error Resume Next
        If bc = 1 Then
            BandViewerForm.ListBox1.AddItem
            BandViewerForm.ListBox1.List(0, 0) = "Point"
            BandViewerForm.ListBox1.List(0, 1) = "STCR"
            BandViewerForm.ListBox1.List(0, 2) = "LTCR"
            BandViewerForm.ListBox1.List(0, 3) = "LTCR + F"
            For Y = 1 To UBound(PointArr)
                BandViewerForm.ListBox1.AddItem
                BandViewerForm.ListBox1.List(Y, 0) = PointArr(Y)
                If LastInspectThick(Y) <> "" And LastInspectThick(Y) <> 0 Then BandViewerForm.ListBox1.List(Y, 1) = Round((LastInspectThick(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - LastInspectDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox1.List(Y, 2) = Round((OrigThickArr(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox1.List(Y, 3) = Round(((OrigThickArr(Y) * 1.1) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
            Next Y
            BandViewerForm.ListBox1.Selected(0) = True
            BandViewerForm.TextBox1 = Round((1 - (MinReading / MaxReading)) * 100, 0) & "%"
        End If
        If bc = 2 Then
            BandViewerForm.ListBox2.AddItem
            BandViewerForm.ListBox2.List(0, 0) = "Point"
            BandViewerForm.ListBox2.List(0, 1) = "STCR"
            BandViewerForm.ListBox2.List(0, 2) = "LTCR"
            BandViewerForm.ListBox2.List(0, 3) = "LTCR + F"
            For Y = 1 To UBound(PointArr)
                BandViewerForm.ListBox2.AddItem
                BandViewerForm.ListBox2.List(Y, 0) = PointArr(Y)
                If LastInspectThick(Y) <> "" And LastInspectThick(Y) <> 0 Then BandViewerForm.ListBox2.List(Y, 1) = Round((LastInspectThick(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - LastInspectDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox2.List(Y, 2) = Round((OrigThickArr(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox2.List(Y, 3) = Round(((OrigThickArr(Y) * 1.1) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
            Next Y
            BandViewerForm.ListBox2.Selected(0) = True
            BandViewerForm.TextBox2 = Round((1 - (MinReading / MaxReading)) * 100, 0) & "%"
        End If
        If bc = 3 Then
            BandViewerForm.ListBox3.AddItem
            BandViewerForm.ListBox3.List(0, 0) = "Point"
            BandViewerForm.ListBox3.List(0, 1) = "STCR"
            BandViewerForm.ListBox3.List(0, 2) = "LTCR"
            BandViewerForm.ListBox3.List(0, 3) = "LTCR + F"
            For Y = 1 To UBound(PointArr)
                BandViewerForm.ListBox3.AddItem
                BandViewerForm.ListBox3.List(Y, 0) = PointArr(Y)
                If LastInspectThick(Y) <> "" And LastInspectThick(Y) <> 0 Then BandViewerForm.ListBox3.List(Y, 1) = Round((LastInspectThick(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - LastInspectDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox3.List(Y, 2) = Round((OrigThickArr(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox3.List(Y, 3) = Round(((OrigThickArr(Y) * 1.1) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
            Next Y
            BandViewerForm.ListBox3.Selected(0) = True
            BandViewerForm.TextBox3 = Round((1 - (MinReading / MaxReading)) * 100, 0) & "%"
        End If
        If bc = 4 Then
            BandViewerForm.ListBox4.AddItem
            BandViewerForm.ListBox4.List(0, 0) = "Point"
            BandViewerForm.ListBox4.List(0, 1) = "STCR"
            BandViewerForm.ListBox4.List(0, 2) = "LTCR"
            BandViewerForm.ListBox4.List(0, 3) = "LTCR + F"
            For Y = 1 To UBound(PointArr)
                BandViewerForm.ListBox4.AddItem
                BandViewerForm.ListBox4.List(Y, 0) = PointArr(Y)
                If LastInspectThick(Y) <> "" And LastInspectThick(Y) <> 0 Then BandViewerForm.ListBox4.List(Y, 1) = Round((LastInspectThick(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - LastInspectDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox4.List(Y, 2) = Round((OrigThickArr(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox4.List(Y, 3) = Round(((OrigThickArr(Y) * 1.1) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
            Next Y
            BandViewerForm.ListBox4.Selected(0) = True
            BandViewerForm.TextBox4 = Round((1 - (MinReading / MaxReading)) * 100, 0) & "%"
        End If
        If bc = 5 Then
            BandViewerForm.ListBox5.AddItem
            BandViewerForm.ListBox5.List(0, 0) = "Point"
            BandViewerForm.ListBox5.List(0, 1) = "STCR"
            BandViewerForm.ListBox5.List(0, 2) = "LTCR"
            BandViewerForm.ListBox5.List(0, 3) = "LTCR + F"
            For Y = 1 To UBound(PointArr)
                BandViewerForm.ListBox5.AddItem
                BandViewerForm.ListBox5.List(Y, 0) = PointArr(Y)
                If LastInspectThick(Y) <> "" And LastInspectThick(Y) <> 0 Then BandViewerForm.ListBox5.List(Y, 1) = Round((LastInspectThick(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - LastInspectDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox5.List(Y, 2) = Round((OrigThickArr(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox5.List(Y, 3) = Round(((OrigThickArr(Y) * 1.1) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
            Next Y
            BandViewerForm.ListBox5.Selected(0) = True
            BandViewerForm.TextBox5 = Round((1 - (MinReading / MaxReading)) * 100, 0) & "%"
        End If
        If bc = 6 Then
            BandViewerForm.ListBox6.AddItem
            BandViewerForm.ListBox6.List(0, 0) = "Point"
            BandViewerForm.ListBox6.List(0, 1) = "STCR"
            BandViewerForm.ListBox6.List(0, 2) = "LTCR"
            BandViewerForm.ListBox6.List(0, 3) = "LTCR + F"
            For Y = 1 To UBound(PointArr)
                BandViewerForm.ListBox6.AddItem
                BandViewerForm.ListBox6.List(Y, 0) = PointArr(Y)
                If LastInspectThick(Y) <> "" And LastInspectThick(Y) <> 0 Then BandViewerForm.ListBox6.List(Y, 1) = Round((LastInspectThick(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - LastInspectDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox6.List(Y, 2) = Round((OrigThickArr(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox6.List(Y, 3) = Round(((OrigThickArr(Y) * 1.1) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
            Next Y
            BandViewerForm.ListBox6.Selected(0) = True
            BandViewerForm.TextBox6 = Round((1 - (MinReading / MaxReading)) * 100, 0) & "%"
        End If
        If bc = 7 Then
            BandViewerForm.ListBox7.AddItem
            BandViewerForm.ListBox7.List(0, 0) = "Point"
            BandViewerForm.ListBox7.List(0, 1) = "STCR"
            BandViewerForm.ListBox7.List(0, 2) = "LTCR"
            BandViewerForm.ListBox7.List(0, 3) = "LTCR + F"
            For Y = 1 To UBound(PointArr)
                BandViewerForm.ListBox7.AddItem
                BandViewerForm.ListBox7.List(Y, 0) = PointArr(Y)
                If LastInspectThick(Y) <> "" And LastInspectThick(Y) <> 0 Then BandViewerForm.ListBox7.List(Y, 1) = Round((LastInspectThick(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - LastInspectDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox7.List(Y, 2) = Round((OrigThickArr(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox7.List(Y, 3) = Round(((OrigThickArr(Y) * 1.1) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
            Next Y
            BandViewerForm.ListBox7.Selected(0) = True
            BandViewerForm.TextBox7 = Round((1 - (MinReading / MaxReading)) * 100, 0) & "%"
        End If
        If bc = 8 Then
            BandViewerForm.ListBox8.AddItem
            BandViewerForm.ListBox8.List(0, 0) = "Point"
            BandViewerForm.ListBox8.List(0, 1) = "STCR"
            BandViewerForm.ListBox8.List(0, 2) = "LTCR"
            BandViewerForm.ListBox8.List(0, 3) = "LTCR + F"
            For Y = 1 To UBound(PointArr)
                BandViewerForm.ListBox8.AddItem
                BandViewerForm.ListBox8.List(Y, 0) = PointArr(Y)
                If LastInspectThick(Y) <> "" And LastInspectThick(Y) <> 0 Then BandViewerForm.ListBox8.List(Y, 1) = Round((LastInspectThick(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - LastInspectDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox8.List(Y, 2) = Round((OrigThickArr(Y) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
                BandViewerForm.ListBox8.List(Y, 3) = Round(((OrigThickArr(Y) * 1.1) - ThickArr(Y)) / ((InspectDateArr(Y) - OrigDateArr(Y)) / 365) * 1000, 1)
            Next Y
            BandViewerForm.ListBox8.Selected(0) = True
            BandViewerForm.TextBox8 = Round((1 - (MinReading / MaxReading)) * 100, 0) & "%"
        End If
        On Error GoTo 0
        
        'expand data to smooth radar chart
        For Y = 1 To UBound(PointArr)
            If Right(PointArr(Y), 1) = "A" Then
            XPointArr(1) = PointArr(Y)
            XThickArr(1) = ThickArr(Y)
            End If
            If Right(PointArr(Y), 1) = "B" Then
            XPointArr(11) = PointArr(Y)
            XThickArr(11) = ThickArr(Y)
            End If
            If Right(PointArr(Y), 1) = "C" Then
            XPointArr(21) = PointArr(Y)
            XThickArr(21) = ThickArr(Y)
            End If
            If Right(PointArr(Y), 1) = "D" Then
            XPointArr(31) = PointArr(Y)
            XThickArr(31) = ThickArr(Y)
            End If
            If Right(PointArr(Y), 1) = "E" Then
            XPointArr(41) = PointArr(Y)
            XThickArr(41) = ThickArr(Y)
            End If
            If Right(PointArr(Y), 1) = "F" Then
            XPointArr(51) = PointArr(Y)
            XThickArr(51) = ThickArr(Y)
            End If
            If Right(PointArr(Y), 1) = "G" Then
            XPointArr(61) = PointArr(Y)
            XThickArr(61) = ThickArr(Y)
            End If
            If Right(PointArr(Y), 1) = "H" Then
            XPointArr(71) = PointArr(Y)
            XThickArr(71) = ThickArr(Y)
            End If
            'asc("Z")-64
        Next Y
        
        
        For Y = 1 To 80
            If XThickArr(Y) <> "" Then
                a = Y
                Do
                    a = a + 1
                    If a = 81 Then Exit Do
                Loop Until XThickArr(a) <> ""
                w = 0
                v = a - Y
                For Z = Y + 1 To a - 1
                    w = w + 1
                    If a <= 80 Then
                        XThickArr(Z) = (((XThickArr(a) - XThickArr(Y)) / (v)) * w) + XThickArr(Y)
                    Else
                        D = 1
                        Do Until XThickArr(D) <> ""
                        D = D + 1
                        Loop
                        XThickArr(Z) = (((XThickArr(D) - XThickArr(Y)) / (v)) * w) + XThickArr(Y)
                    End If
                Next Z
                Y = a - 1
            End If
        Next Y
        For Y = 1 To 80
            If XThickArr(Y) <> "" Then
                XThickArr(Y) = 10 - XThickArr(Y)
                If XThickArr(Y) > MaxThick Then MaxThick = XThickArr(Y)
            End If
        Next Y
        For Y = 1 To 80
            MaxThickArr(Y) = MaxThick
        Next Y
        
        'make charts
        Set BandChart = aSh.ChartObjects.Add(10, 10, 354, 210).Chart
        With BandChart
            .ChartType = xlRadarFilled
            .SeriesCollection.NewSeries
            .SeriesCollection(1).Name = "=""OD"""
            .SeriesCollection(1).Values = MaxThickArr
            .SeriesCollection.NewSeries
            .SeriesCollection(2).Name = "=""ID"""
            .SeriesCollection(2).Values = XThickArr
            .SeriesCollection(2).XValues = XPointArr
            With .SeriesCollection(2).Format.Fill
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = 0
                .Solid
            End With
            .Axes(xlValue).MaximumScale = MaxThick
            .Axes(xlValue).MinimumScale = MaxThick - 0.6
            .Axes(xlValue).TickLabelPosition = xlNone
            With .Axes(xlValue).Format.Line
                .Visible = msoTrue
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                .ForeColor.TintAndShade = 0
                .ForeColor.Brightness = -0.150000006
                .Transparency = 0
            End With
            .SetElement (msoElementChartTitleAboveChart)
            .ChartTitle.Text = "Band " & Left(Right(aSh.Cells(r - 1, 1), 2), 1)
            .SetElement (msoElementLegendNone)
            '.PlotArea.Format.Line.Visible = msoFalse
            .ChartArea.Border.LineStyle = xlNone

        End With
        If aSh.Cells(r, 1) = "" Then Exit Do
Loop Until Left(aSh.Cells(r - 1, 1), Len(aSh.Cells(Target.Row, 1)) - 2) <> Left(aSh.Cells(r, 1), Len(aSh.Cells(r, 1)) - 2)

'save charts as images as load into userform

For Each chobj In aSh.ChartObjects
    c = c + 1
    chobj.Activate 'get random errors in 2010 if it is not activated
    TempName = FSO.GetTempName
    TempName = Left(TempName, Len(TempName) - 4) & ".gif"
    'FName = Application.DefaultFilePath & "\temp.gif"
    FName = ActiveWorkbook.Path & "\temp.gif"

    ActiveChart.Export Filename:=FName, FilterName:="GIF"
    
    Select Case c
    Case 1
    BandViewerForm.Image1.Picture = LoadPicture(FName)
    Case 2
    BandViewerForm.Image2.Picture = LoadPicture(FName)
    Case 3
    BandViewerForm.Image3.Picture = LoadPicture(FName)
    Case 4
    BandViewerForm.Image4.Picture = LoadPicture(FName)
    Case 5
    BandViewerForm.Image5.Picture = LoadPicture(FName)
    Case 6
    BandViewerForm.Image6.Picture = LoadPicture(FName)
    Case 7
    BandViewerForm.Image7.Picture = LoadPicture(FName)
    Case 8
    BandViewerForm.Image8.Picture = LoadPicture(FName)
    End Select
    
    chobj.Delete
    Kill FName
    
Next

BandViewerForm.Show vbModeless
aSh.Protect "Dh1986"
End Sub

