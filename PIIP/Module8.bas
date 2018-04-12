Attribute VB_Name = "Module8"
Sub GuessPipeSchedule(Target As Range)

Dim aSh As Worksheet
Dim bg As Worksheet
Dim BandSum As Double
Dim BandAvg As Double
Dim BandNom As Double
Dim BandOD As Double
Dim BandGuessedSch As Variant
Dim BandActualSch As Variant


Set aSh = ActiveSheet
Set bg = Sheets("ScheduleGuesser")

r = Target.Row
'loop to find first row of Band clicked
If Target.Row > 2 Then
    Do Until Left(Right(aSh.Cells(Target.Row, 1), 2), 1) <> Left(Right(aSh.Cells(r, 1), 2), 1)
        r = r - 1
    Loop
End If
r = r + 1

X = 0
s = r

'loop until last row of Band
Do
    If aSh.Cells(r, 12) <> "" Then
        BandSum = BandSum + aSh.Cells(r, 12)
        X = X + 1
    End If
    r = r + 1
Loop Until Left(Right(aSh.Cells(r, 1), 2), 1) <> Left(Right(aSh.Cells(s, 1), 2), 1) 'until next band is detected

BandAvg = BandSum / X
BandNom = aSh.Cells(s, 5)
BandOD = aSh.Cells(s, 27)

bg.Cells(1, 2) = BandOD
bg.Cells(2, 2) = BandNom
BandActualSch = bg.Cells(4, 2)

bg.Cells(2, 2) = BandAvg
BandGuessedSch = bg.Cells(4, 2)

With ScheduleGuesserForm

    .TML = "TML " & Left(aSh.Cells(s, 1), Len(aSh.Cells(s, 1)) - 2) & "        Band " & Left(Right(aSh.Cells(s, 1), 2), 1)
    .NomSch = BandActualSch
    .ThickSch = BandGuessedSch
    
    'For w = 0 To 2
        '.ListBox1.AddItem
        'For Z = 0 To 13
                '.ListBox1.List(w, Z) = bg.Cells(6 + w, 1 + Z)
        'Next Z
    'Next w
    '.ListBox1.Selected(0) = True
    If .NomSch = .ThickSch Then .WarningLabel.Visible = False
    .Show vbModeless

End With

End Sub
