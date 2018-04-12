Attribute VB_Name = "Module2"
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Function PixPerPtX()
Dim hdc As Long


hdc = GetDC(0)

PixPerInchX = GetDeviceCaps(hdc, LOGPIXELSX)

'there are 72 points per inch
PixPerPtX = PixPerInchX / 72

ReleaseDC 0, hdc
End Function


Public Function SortArray(ByRef TheArray() As Variant, NumBlanks)

For X = 1 To UBound(TheArray)
    If TheArray(X) = "" Then TheArray(X) = ""
Next X

sorted = False
Do While Not sorted
    sorted = True
    For X = 1 To UBound(TheArray) - 1
        If TheArray(X) > TheArray(X + 1) Then
            Temp = TheArray(X + 1)
            TheArray(X + 1) = TheArray(X)
            TheArray(X) = Temp
            sorted = False
        End If
    Next X
Loop

NumBlanks = 0
For X = 1 To UBound(TheArray)
    If TheArray(X) = "" Then NumBlanks = NumBlanks + 1
Next X


End Function

Public Function MeanFunction(MeanArray() As Variant, Optional Num As Integer)

Sum = 0
Num = 0
For X = 1 To UBound(MeanArray)

    If Application.WorksheetFunction.IsNumber(MeanArray(X)) = True Then
        Sum = Sum + MeanArray(X)
        Num = Num + 1
    End If

Next X

If Num > 1 Then MeanFunction = Sum / Num
'MsgBox MeanFunction

End Function

Public Function StDevFunction(StDevArray() As Variant)

avg = MeanFunction(StDevArray)
StdDevN = 0: Num = 0


For X = 1 To UBound(StDevArray)

    If Application.WorksheetFunction.IsNumber(StDevArray(X)) = True Then
        StdDevN = StdDevN + ((StDevArray(X) - avg) ^ 2)
        Num = Num + 1
    End If

Next X

If Num > 1 Then StDevFunction = (StdDevN / (Num - 1)) ^ 0.5
'MsgBox StDevFunction


End Function
Public Function StDevPFunction(StDevArray() As Variant)

avg = MeanFunction(StDevArray)
StdDevN = 0: Num = 0


For X = 1 To UBound(StDevArray)

    If Application.WorksheetFunction.IsNumber(StDevArray(X)) = True Then
        StdDevN = StdDevN + ((StDevArray(X) - avg) ^ 2)
        Num = Num + 1
    End If

Next X

If Num > 1 Then StDevPFunction = (StdDevN / (Num)) ^ 0.5
'MsgBox StDevPFunction

End Function

Function SEVData(Xarray As Variant, Sl, yIntercept)

Dim Rarr1() As Variant
Dim Rarr2() As Variant
Dim Rarr3() As Variant
Dim MRarr() As Variant
Dim yarr() As Variant
Dim CalcArrayTemp() As Variant
Dim CalcArray() As Variant

ReDim Rarr1(1 To UBound(Xarray))
ReDim Rarr2(1 To UBound(Xarray))
ReDim Rarr3(1 To UBound(Xarray))
ReDim MRarr(1 To UBound(Xarray))
ReDim yarr(1 To UBound(Xarray))

Application.DisplayAlerts = False

Sheets.Add

With ActiveSheet

CalcArrayTemp = Xarray

'RemoveBlanksFromArray PPlotArray
SortArray CalcArrayTemp, nb
ReDim CalcArray(UBound(CalcArrayTemp) - nb)

For X = 1 To UBound(CalcArrayTemp) - nb
    CalcArray(X) = CalcArrayTemp(X)
Next X
'raw data
N = UBound(CalcArray)

For X = 1 To UBound(CalcArray)
    .Cells(X, 1) = CalcArray(X)
Next X

'rank
.Cells(1, 2).FormulaR1C1 = _
    "=RANK(RC[-1],R1C1:R" & N & "C1,1)+(COUNT(R1C1:R" & N & "C1)+1-RANK(RC[-1],R1C1:R" & N & "C1,0)-RANK(RC[-1],R1C1:R" & N & "C1,1))/2"
    .Range(.Cells(1, 2), .Cells(1, 2)).AutoFill Destination:=.Range(.Cells(1, 2), .Cells(N, 2)), Type:=xlFillDefault

'median rank (Bernard)
.Cells(1, 3).FormulaR1C1 = _
    "=((RC[-1]-0.3)/(" & N & "+0.4))"
    .Range(.Cells(1, 3), .Cells(1, 3)).AutoFill Destination:=.Range(.Cells(1, 3), .Cells(N, 3)), Type:=xlFillDefault

.Cells(1, 4).FormulaR1C1 = _
    "=ln(-ln(1-RC[-1]))"
    .Range(.Cells(1, 4), .Cells(1, 4)).AutoFill Destination:=.Range(.Cells(1, 4), .Cells(N, 4)), Type:=xlFillDefault
    
Sl = WorksheetFunction.Slope(Range(.Cells(1, 4), .Cells(N, 4)), Range(.Cells(1, 1), .Cells(N, 1))) 'slope
yIntercept = WorksheetFunction.Intercept(Range(.Cells(1, 4), .Cells(N, 4)), Range(.Cells(1, 1), .Cells(N, 1))) 'Intercept

.Delete

End With

    
End Function

Sub removeDuplicates(ByRef arrName() As Variant)
    'note: sorts array with base 1
    Dim i As Long, tempArr() As Variant: 'ReDim tempArr(1 To UBound(arrName))
    Dim D As New Dictionary, N As Long
    
    N = 1
    For i = 1 To UBound(arrName)
        If Not D.Exists(arrName(i)) Then 'searching the dictionary to see if item exists
            D.Add arrName(i), arrName(i) 'if item doesn't exist, add to dictionary
            ReDim Preserve tempArr(N)
            tempArr(N) = arrName(i): N = N + 1 'if item doesn't exist, add to temparr()
        End If
    Next
    arrName = tempArr
  
End Sub

Function WorksheetExists(ByVal WorksheetName As String) As Boolean

On Error Resume Next
WorksheetExists = (Sheets(WorksheetName).Name <> "")
On Error GoTo 0

End Function
