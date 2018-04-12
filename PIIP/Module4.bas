Attribute VB_Name = "Module4"
Dim XScaleArray(1 To 17) As Double
Dim YScaleArray(1 To 17) As Double
Dim YScaleArray2(1 To 17) As Double
Dim ScaleArray(1 To 17) As Double
Dim XAxisArray(1 To 17) As Double
Dim Zp(1 To 17) As Double
Dim VarXp(1 To 17) As Double
Dim XpL(1 To 17) As Double
Dim XpU(1 To 17) As Double
Dim N As Integer
Public SName1 As String
Public PScale As Double
Public NewChartName As String
Public PlotType
Dim D1Empty As Boolean


'ScaleArray(x) 'used for calculations - includes 2.5 & 97.5
'ScaleArray2(x) 'used for gridlines - excludes 2.5 & 97.5
'YScaleArray(x) = Round(WorksheetFunction.NormInv(ScaleArray(x), muS, SigmaS), 4)
'YScaleArray2(x) = WorksheetFunction.NormInv(ScaleArray(x), mu, sigmap)
'YScaleArray3(x) = Round(WorksheetFunction.NormInv(ScaleArray2(x), muS, SigmaS), 4)

Function MakeExcelPPlot(AlignedDataSet1() As Variant, ChartName As String)

Application.ScreenUpdating = False
Application.ShowChartTipValues = False 'y values are false so they shouldn't be shown

SName = Year(Now)
NewChartName = ChartName
PlotType = "Norm"

If WorksheetExists(ChartName) = True Then
    MsgBox "You already have a plot created for " & ChartName & ". Either rename or delete existing one to make a new chart with this name.", vbCritical
    Exit Function
End If


'check to see if all datasets are blank (or have less than 2 UNIQUE data points)
D1Empty = False
If UBound(AlignedDataSet1) < 2 Then D1Empty = True

If D1Empty = True Then
    MsgBox "Dataset is empty or has less than 2 unique data points. Cannot produce Probability Plot.", vbCritical
    Exit Function
End If

'Choose scale to use
PPlotScaleChooser AlignedDataSet1, muS, SigmaS, MinXScale, MaxXScale
'Make Chart Skeleton
PPlotChartSkeleton muS, SigmaS, YScaleArray, MinXScale, MaxXScale

   


'Calculate chart Data
PPlotData AlignedDataSet1, YScaleArray, YScaleArray2, XpL, XpU, muS, SigmaS
'Add Chart Series
AddPPlotSeries YScaleArray, YScaleArray2, XpL, XpU, SName
        
Charts(NewChartName).Activate

Call PPlotLegend

ActiveChart.ProtectData = True
Application.ScreenUpdating = True

End Function


Function PPlotScaleChooser(DS1() As Variant, muS, SigmaS, MinXScale, MaxXScale)

'''Note: Case where data points outside confidence intervals not accounted for.

Dim MinScale(1 To 3) As Double
Dim MaxScale(1 To 3) As Double

'1. Receive three data sets
'2. Determine which yscale to use
'3. Calculate the data sets with respect to the scale being used
    'NormInv using mean and StDev of datapoints/scale being used
    'YScaleArray - used for line, upper limit and lower limit -
    'scale accordingly...
'4. Determine XMin and XMax
'5. Plot all data sets

N1 = 0

SortArray DS1, NumBlanks1
N1 = UBound(DS1) - NumBlanks1
DS1Min = DS1(1): DS1Max = DS1(N1)

MinP = (1 - 0.3) / (N1 + 0.4)

P1 = 1: P2 = P1 / 10
Do Until P2 = 0.0000000000001
    If MinP <= P1 And MinP >= P2 Then
        PScale = P2
        Exit Do
    End If
    P1 = P2
    P2 = P2 / 10
Loop

If PScale > 0.01 Then PScale = 0.01
If NewScale <> 0 Then PScale = NewScale

'Probability Scale
If PlotType = "Norm" Then
    MinScale(1) = WorksheetFunction.NormInv(PScale, MeanFunction(DS1), StDevFunction(DS1))
    
    MaxScale(1) = WorksheetFunction.NormInv(1 - PScale, MeanFunction(DS1), StDevFunction(DS1))
ElseIf PlotType = "SEV" Then
    SEVData DS1, Sl, yIntercept
    MinScale(1) = (WorksheetFunction.Ln(-WorksheetFunction.Ln(1 - PScale)) - yIntercept) / Sl
    MaxScale(1) = (WorksheetFunction.Ln(-WorksheetFunction.Ln(1 - (1 - PScale))) - yIntercept) / Sl
End If

    muS = MeanFunction(DS1)
    SigmaS = StDevFunction(DS1)

'X-Axis Scale

Za = 1.96
varu1 = (StDevFunction(DS1) ^ 2) / N1
'VarSigma1 = (StDevFunction(DS1) ^ 2) * (1 - ((2 * (Exp(WorksheetFunction.GammaLn(N1 / 2)) ^ 2)) / ((N1 - 1) * (Exp(WorksheetFunction.GammaLn((N1 - 1) / 2)) ^ 2))))
If N1 >= 101 Then
    VarSigma1 = (StDevFunction(DS1) ^ 2) / (2 * N1)
Else
    VarSigma1 = (StDevFunction(DS1) ^ 2) * (1 - ((2 * (Exp(WorksheetFunction.GammaLn(N1 / 2)) ^ 2)) / ((N1 - 1) * (Exp(WorksheetFunction.GammaLn((N1 - 1) / 2)) ^ 2))))
End If

Zp1 = (MinScale(1) - MeanFunction(DS1)) / StDevFunction(DS1)
VarXp1 = varu1 + ((Zp1 ^ 2) * VarSigma1)
XpL1 = MinScale(1) - (Za * (VarXp1 ^ 0.5))
Zp1 = (MaxScale(1) - MeanFunction(DS1)) / StDevFunction(DS1)
VarXp1 = varu1 + ((Zp1 ^ 2) * VarSigma1)
XpU1 = MaxScale(1) + (Za * (VarXp1 ^ 0.5))

MinXScale = XpL1
MaxXScale = XpU1

If MinXScale > DS1Min Then
    MinXScale = DS1Min
End If
If MaxXScale < DS1Max Then
    MaxXScale = DS1Max
End If

End Function


Sub PPlotData(PPlotArray() As Variant, YScaleArray, YScaleArray2, XpL, XpU, muS, SigmaS)

'Dim CalcArrayTemp(1 To UBound(PPlotArray)) As Variant
Dim CalcArrayTemp() As Variant
Dim CalcArray() As Variant
Dim cws As Worksheet
Dim CalcSheetExists As Boolean

'adds CalcSheet if it doesn't already exist
CalcSheetExists = False
For X = 1 To ActiveWorkbook.Worksheets.Count
    If ActiveWorkbook.Worksheets(X).Name = "CalcSheet" Then
        CalcSheetExists = True
    End If
Next X
If CalcSheetExists = False Then
   ActiveWorkbook.Worksheets.Add(After:=Worksheets("HomePage")).Name = "CalcSheet"
   ActiveWorkbook.Sheets("CalcSheet").Visible = 0
End If

 
Set cws = ActiveWorkbook.Sheets("CalcSheet")

CalcArrayTemp = PPlotArray

c = cws.UsedRange.Columns.Count
If c = 1 Then c = 0

NumCalcSheets = 1
If c >= 254 Then
    For X = 1 To ActiveWorkbook.Worksheets.Count
        If Left(ActiveWorkbook.Sheets(X).Name, 9) = "CalcSheet" Then
            NumCalcSheets = NumCalcSheets + 1
        End If
    Next X
    ActiveWorkbook.Sheets("CalcSheet").Name = "CalcSheet" & NumCalcSheets
    ActiveWorkbook.Worksheets.Add(Before:=Sheets("CalcSheet2")).Name = "CalcSheet"
    c = 0
End If

Set cws = ActiveWorkbook.Sheets("CalcSheet")

'RemoveBlanksFromArray PPlotArray
SortArray CalcArrayTemp, nb
ReDim CalcArray(UBound(CalcArrayTemp) - nb)

For X = 1 To UBound(CalcArrayTemp) - nb
    CalcArray(X) = CalcArrayTemp(X)
Next X
'raw data
For X = 1 To UBound(CalcArray)
    cws.Cells(X, c + 1).Value = CalcArray(X)
Next X

N = UBound(CalcArray)
If N <= 1 Then Exit Sub

'worksheet has to be activated for the formulas to work...
cws.Activate

'''''''''''''''''''''''''Worksheet Calculations
'rank
cws.Cells(1, c + 2).FormulaR1C1 = _
    "=RANK(RC[-1],R1C" & c + 1 & ":R" & N & "C" & c + 1 & ",1)+(COUNT(R1C" & c + 1 & ":R" & N & "C" & c + 1 & ")+1-RANK(RC[-1],R1C" & c + 1 & ":R" & N & "C" & c + 1 & ",0)-RANK(RC[-1],R1C" & c + 1 & ":R" & N & "C" & c + 1 & ",1))/2"
cws.Range(cws.Cells(1, c + 2), cws.Cells(1, c + 2)).AutoFill Destination:=cws.Range(cws.Cells(1, c + 2), cws.Cells(N, c + 2)), Type:=xlFillDefault

'median rank (Bernard)
cws.Cells(1, c + 3).FormulaR1C1 = _
    "=((RC[-1]-0.3)/(" & N & "+0.4))"
cws.Range(cws.Cells(1, c + 3), cws.Cells(1, c + 3)).AutoFill Destination:=cws.Range(cws.Cells(1, c + 3), cws.Cells(N, c + 3)), Type:=xlFillDefault

If PlotType = "Norm" Then
    'Inverse CDF
    cws.Cells(1, c + 4).FormulaR1C1 = _
        "=NORMINV(RC[-1]," & muS & "," & SigmaS & ")"
    cws.Range(cws.Cells(1, c + 4), cws.Cells(1, c + 4)).AutoFill Destination:=cws.Range(cws.Cells(1, c + 4), cws.Cells(N, c + 4)), Type:=xlFillDefault
ElseIf PlotType = "SEV" Then
    'ln(-ln(1-p))
    cws.Cells(1, c + 4).FormulaR1C1 = _
        "=ln(-ln(1-RC[-1]))"
    cws.Range(cws.Cells(1, c + 4), cws.Cells(1, c + 4)).AutoFill Destination:=cws.Range(cws.Cells(1, c + 4), cws.Cells(N, c + 4)), Type:=xlFillDefault
    Sl = WorksheetFunction.Slope(Range(cws.Cells(1, c + 4), cws.Cells(N, c + 4)), Range(cws.Cells(1, c + 1), cws.Cells(N, c + 1))) 'slope
    InterC = WorksheetFunction.Intercept(Range(cws.Cells(1, c + 4), cws.Cells(N, c + 4)), Range(cws.Cells(1, c + 1), cws.Cells(N, c + 1))) 'Intercept
End If


With Range(cws.Cells(1, c + 1), cws.Cells(N, c + 4))
    .Copy
    .PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End With

Range(cws.Cells(1, c + 2), cws.Cells(N, c + 2)).EntireColumn.Delete Shift:=xlToLeft



'''''''''''''''''''''''''''''''''''''''''''''''''
mu = WorksheetFunction.Average(Range(cws.Cells(1, c + 1), cws.Cells(N, c + 1)))
sigma = WorksheetFunction.StDev(Range(cws.Cells(1, c + 1), cws.Cells(N, c + 1)))
sigmap = WorksheetFunction.StDevP(Range(cws.Cells(1, c + 1), cws.Cells(N, c + 1)))

varu = (sigma ^ 2) / N
'VarSigma = (sigmap ^ 2) / (2 * N) - this worked for the other method
If N <= 101 Then
    VarSNumerator = 2 * (Exp(WorksheetFunction.GammaLn(N / 2)) ^ 2)
    VarSDenominator = (N - 1) * (Exp(WorksheetFunction.GammaLn((N - 1) / 2)) ^ 2)
    VarSigma = (sigma ^ 2) * (1 - (VarSNumerator / VarSDenominator))
Else
    'VarSigma = sigma ^ 2
    VarSigma = (sigma ^ 2) / (2 * N)
End If

Za = 1.96

If PlotType = "Norm" Then
    For X = 1 To UBound(ScaleArray)
        YScaleArray2(X) = WorksheetFunction.NormInv(ScaleArray(X), mu, sigma)
        Zp(X) = (YScaleArray2(X) - mu) / sigma
        VarXp(X) = varu + ((Zp(X) ^ 2) * VarSigma)
        XpL(X) = Round(YScaleArray2(X) - (Za * (VarXp(X) ^ 0.5)), 4)
        XpU(X) = Round(YScaleArray2(X) + (Za * (VarXp(X) ^ 0.5)), 4)
    Next X
ElseIf PlotType = "SEV" Then
    For X = 1 To UBound(ScaleArray)
        YScaleArray2(X) = (WorksheetFunction.Ln(-WorksheetFunction.Ln(1 - ScaleArray(X))) - InterC) / Sl
        Zp(X) = (YScaleArray2(X) - mu) / sigma
        VarXp(X) = varu + ((Zp(X) ^ 2) * VarSigma)
        XpL(X) = Round(YScaleArray2(X) - (Za * (VarXp(X) ^ 0.5)), 4)
        XpU(X) = Round(YScaleArray2(X) + (Za * (VarXp(X) ^ 0.5)), 4)
    Next X
End If

For X = 1 To UBound(ScaleArray)
    YScaleArray2(X) = Round(YScaleArray2(X), 4)
Next X


End Sub

Function PPlotChartSkeleton(muS, SigmaS, YScaleArray, MinXScale, MaxXScale)

Dim ChartExists As Boolean
Dim ScaleArray2(1 To 15) As Double
Dim YScaleArray3(1 To 15) As Double
Dim XAxisArray2(1 To 15) As Double


ActiveWorkbook.Charts.Add(After:=Sheets("Homepage")).Name = NewChartName
ActiveWorkbook.Sheets(NewChartName).Tab.Color = 12611584

ActiveChart.DisplayBlanksAs = xlZero 'avoids the Unable to Set XValues.... error (can avoid by making sure no cells are selected when adding chart (level 1 analysis summary))

If ActiveChart.SeriesCollection.Count > 1 Then
    Do Until ActiveChart.SeriesCollection.Count = 1
        ActiveChart.SeriesCollection(1).Delete
    Loop
ElseIf ActiveChart.SeriesCollection.Count = 0 Then
    ActiveChart.SeriesCollection.NewSeries
End If

If PScale > 0.01 Then
    ScaleArray(1) = 0.01
    ScaleArray(15) = 0.99
    ScaleArray2(1) = 0.01
    ScaleArray2(15) = 0.99
Else
    ScaleArray(1) = PScale
    ScaleArray(17) = 1 - PScale
    ScaleArray2(1) = PScale
    ScaleArray2(15) = 1 - PScale
End If

'used for calculations
ScaleArray(2) = 0.01: ScaleArray(3) = 0.025: ScaleArray(4) = 0.05: ScaleArray(5) = 0.1
ScaleArray(6) = 0.2: ScaleArray(7) = 0.3: ScaleArray(8) = 0.4
ScaleArray(9) = 0.5: ScaleArray(10) = 0.6: ScaleArray(11) = 0.7
ScaleArray(12) = 0.8: ScaleArray(13) = 0.9: ScaleArray(14) = 0.95
ScaleArray(15) = 0.975: ScaleArray(16) = 0.99

'used for scale data series
If PlotType = "Norm" Then
    '(normal)
    ScaleArray2(2) = 0.01: ScaleArray2(3) = 0.05: ScaleArray2(4) = 0.1
    ScaleArray2(5) = 0.2: ScaleArray2(6) = 0.3: ScaleArray2(7) = 0.4
    ScaleArray2(8) = 0.5: ScaleArray2(9) = 0.6: ScaleArray2(10) = 0.7
    ScaleArray2(11) = 0.8: ScaleArray2(12) = 0.9: ScaleArray2(13) = 0.95
    ScaleArray2(14) = 0.99
ElseIf PlotType = "SEV" Then
    'SEV
    ScaleArray2(2) = 0.01: ScaleArray2(3) = 0.02: ScaleArray2(4) = 0.03
    ScaleArray2(5) = 0.05: ScaleArray2(6) = 0.1: ScaleArray2(7) = 0.2
    ScaleArray2(8) = 0.3: ScaleArray2(9) = 0.4: ScaleArray2(10) = 0.5
    ScaleArray2(11) = 0.6: ScaleArray2(12) = 0.7: ScaleArray2(13) = 0.8
    ScaleArray2(14) = 0.9
End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If PlotType = "Norm" Then
    For X = 1 To UBound(ScaleArray)
        YScaleArray(X) = Round(WorksheetFunction.NormInv(ScaleArray(X), muS, SigmaS), 4)
    Next X
    For X = 1 To UBound(ScaleArray2)
        YScaleArray3(X) = Round(WorksheetFunction.NormInv(ScaleArray2(X), muS, SigmaS), 4)
    Next X
ElseIf PlotType = "SEV" Then
    For X = 1 To UBound(ScaleArray)
        YScaleArray(X) = Round(WorksheetFunction.Ln(-WorksheetFunction.Ln(1 - (ScaleArray(X)))), 4)
    Next X
    For X = 1 To UBound(ScaleArray2)
        YScaleArray3(X) = Round(WorksheetFunction.Ln(-WorksheetFunction.Ln(1 - (ScaleArray2(X)))), 4)
    Next X
End If


With ActiveChart
    .ChartType = xlXYScatter
    .HasLegend = True
    .Legend.Position = xlBottom
    MaxYScale = YScaleArray(UBound(YScaleArray))
    MinYScale = YScaleArray(LBound(YScaleArray))
    .Axes(xlCategory).MaximumScale = Round(MaxXScale, 3) + 0.001
    .Axes(xlCategory).MinimumScale = Round(MinXScale, 3) - 0.001
    .Axes(xlValue).MaximumScale = MaxYScale
    .Axes(xlValue).MinimumScale = MinYScale
    For X = 1 To UBound(ScaleArray) '- 2
        XAxisArray(X) = Round(MinXScale, 3) - 0.001
    Next X
    For X = 1 To UBound(ScaleArray2) '- 2
        XAxisArray2(X) = Round(MinXScale, 3) - 0.001
    Next X
    
    'Probability Scale Series
    With .SeriesCollection(1)
        .XValues = XAxisArray2
        .Values = YScaleArray3
        .Name = "Scale"
        .MarkerStyle = xlNone
        .ErrorBar Direction:=xlX, Include:= _
            xlPlusValues, Type:=xlCustom, Amount:="={" & (Round(MaxXScale, 3) + 0.001) - (Round(MinXScale, 3) - 0.001) & "}"
        .ErrorBars.Border.LineStyle = xlDot
        .ErrorBars.Border.ColorIndex = 15
        .ErrorBars.Border.Weight = xlHairline
        .ErrorBars.EndStyle = xlNoCap
        .ApplyDataLabels AutoText:=True, LegendKey:= _
            False, ShowSeriesName:=False, ShowCategoryName:=False, ShowValue:=True, _
            ShowPercentage:=False, ShowBubbleSize:=False
        .DataLabels.HorizontalAlignment = xlCenter
        .DataLabels.VerticalAlignment = xlCenter
        .DataLabels.Position = xlLabelPositionLeft
        .DataLabels.Orientation = xlHorizontal
    End With
    For X = 1 To .SeriesCollection(.SeriesCollection.Count).Points.Count
        .SeriesCollection(.SeriesCollection.Count).Points(X).DataLabel.Characters.Text = ScaleArray2(X) * 100
    Next X
    .SeriesCollection(.SeriesCollection.Count).Points(1).DataLabel.Delete 'necessary for 2010

    With .Axes(xlValue)
        .Border.Weight = xlHairline
        .Border.LineStyle = xlAutomatic
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
        .HasMajorGridlines = False
        .HasMinorGridlines = False
        .CrossesAt = YScaleArray(LBound(YScaleArray))
    End With
    With .Axes(xlCategory)
        .HasMajorGridlines = False
        .HasMinorGridlines = True
        .MinorGridlines.Border.ColorIndex = 15
        .MinorUnit = .MajorUnit / 2
        .CrossesAt = Round(MinXScale, 3) - 0.001
        .TickLabels.NumberFormat = "0.000"
    End With
    With .PlotArea
        .Border.ColorIndex = 16
        .Border.Weight = xlThin
        .Border.LineStyle = xlContinuous
        .Interior.ColorIndex = 2
    End With
    With .Axes(xlCategory).MinorGridlines.Border
        .LineStyle = xlDot
    End With
    With .ChartArea
        .Fill.Visible = True
        .Fill.ForeColor.SchemeColor = 35
        .Fill.OneColorGradient Style:=msoGradientVertical, Variant:=1, _
        Degree:=1
    End With
    With .PlotArea.Border
        .ColorIndex = 57
        .Weight = xlThin
        .LineStyle = xlContinuous
    End With
    .HasTitle = True
    If PlotType = "SEV" Then
        .ChartTitle.Characters.Text = "Probability Plot" & Chr(10) & "SEV - 95% CI"
    Else
        .ChartTitle.Characters.Text = "Probability Plot" & Chr(10) & "Normal - 95% CI"
    End If
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Data"
    .ChartArea.AutoScaleFont = False
End With

End Function

Function AddPPlotSeries(YScaleArray, YScaleArray2, XpL, XpU, SName)

Dim cws As Worksheet
Dim ADArrTemp() As Variant
Dim ADArr() As Variant
Dim SeriesColorArr(1 To 3) As Long

Set cws = ActiveWorkbook.Sheets("CalcSheet")
SeriesColorArr(1) = 3 'red

c = cws.UsedRange.Columns.Count
If c = 1 Then c = 0

ActiveWorkbook.Charts(NewChartName).Activate

With ActiveChart
    For X = 1 To .SeriesCollection.Count
        If Right(.SeriesCollection(X).Name, 4) = "Data" Then
            NumExistingSeries = NumExistingSeries + 1
        End If
    Next X

    'Point Series
    .SeriesCollection.NewSeries
    With .SeriesCollection(.SeriesCollection.Count)
        .ChartType = xlXYScatter
        .XValues = "=CalcSheet!R1C" & c - 2 & ":R" & N & "C" & c - 2
        .Values = "=CalcSheet!R1C" & c & ":R" & N & "C" & c
        .Name = SName & "Data"
        .MarkerBackgroundColorIndex = SeriesColorArr(NumExistingSeries + 1)
        .MarkerForegroundColorIndex = SeriesColorArr(NumExistingSeries + 1)
        .MarkerStyle = xlCircle
    End With
    
    'Line Series
    .SeriesCollection.NewSeries
    With .SeriesCollection(.SeriesCollection.Count)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = YScaleArray2
        .Values = YScaleArray
        .Name = SName & "Line"
        .Border.ColorIndex = SeriesColorArr(NumExistingSeries + 1)
    End With
    
    'Lower Bound Series
    .SeriesCollection.NewSeries
    With .SeriesCollection(.SeriesCollection.Count)
        .ChartType = xlXYScatterSmoothNoMarkers
        .XValues = XpL
        .Values = YScaleArray
        .Name = SName & "LBound"
        .Border.ColorIndex = SeriesColorArr(NumExistingSeries + 1)
    End With
    
    'Upper Bound Series
    .SeriesCollection.NewSeries
    With .SeriesCollection(.SeriesCollection.Count)
        .ChartType = xlXYScatterSmoothNoMarkers
        .XValues = XpU
        .Values = YScaleArray
        .Name = SName & "UBound"
        .Border.ColorIndex = SeriesColorArr(NumExistingSeries + 1)
    End With
    
    With .Axes(xlValue)
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
    End With
    
End With


End Function



Public Function GetChartRange(cht, series, XRange As String, YRange As String, WSName As String)
'   cht: A Chart object
'   series: Integer representing the Series
'   ValOrX: String, either "values" or "xvalues"

   Dim sf As String
   Dim CommaCnt As Integer
   Dim ExclCnt As Integer
   Dim Commas() As Integer
   Dim Excl() As Integer
   Dim ListSep As String * 1
   Dim Temp As String
   
   If series = "" Then Exit Function
   
   On Error Resume Next
   
'   Get the SERIES formula
   sf = Charts(cht).SeriesCollection(series).Formula
   Wpos = InStr(1, sf, "CalcSheet")
   c = 0
   Do Until X = "'" Or X = "!" 'this will accomodate both excel 2003 and 2010
        X = Mid(sf, Wpos + c, 1)
        c = c + 1
    Loop
   WSName = Mid(sf, Wpos, c - 1)
   
'   Check for noncontiguous ranges by counting commas
'   Also, store the character position of the commas
   CommaCnt = 0
   ListSep = Application.International(xlListSeparator)
   For i = 1 To Len(sf)
       If Mid(sf, i, 1) = ListSep Then
           CommaCnt = CommaCnt + 1
           ReDim Preserve Commas(CommaCnt)
           Commas(CommaCnt) = i
       End If
       If Mid(sf, i, 1) = "!" Then
            ExclCnt = ExclCnt + 1
            ReDim Preserve Excl(ExclCnt)
            Excl(ExclCnt) = i
        End If
       
   Next i
   If CommaCnt > 3 Then Exit Function
   
    'Text between 1st and 2nd commas in SERIES Formula
    Temp = Mid(sf, Excl(1) + 1, Commas(2) - Excl(1) - 1)
    XRange = Temp
    'Text between the 2nd and 3rd commas in SERIES Formula
    Temp = Mid(sf, Excl(2) + 1, Commas(3) - Excl(2) - 1)
    YRange = Temp
    
    'MsgBox "stop"

End Function

Sub PPlotLegend()

ActiveWorkbook.Charts(NewChartName).Activate

With ActiveChart
    If Application.Version = "11.0" Then
        For X = .SeriesCollection.Count To 1 Step -1
            If Right(.SeriesCollection(X).Name, 4) <> "Data" Then
                .Legend.LegendEntries(X).Delete
            End If
        Next X
    ElseIf Application.Version = "14.0" Then
        For X = .Legend.LegendEntries.Count To 1 Step -1
            If .Legend.LegendEntries(X).LegendKey.MarkerStyle <> 8 Then .Legend.LegendEntries(X).Delete
        Next X
    End If
End With

End Sub
