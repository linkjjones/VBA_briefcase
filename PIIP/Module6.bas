Attribute VB_Name = "Module6"
Public XTwb As Workbook
Public XTsh As Worksheet
Public CancelTransfer As Boolean

Sub XToolInput()

Dim XTArr(1 To 10000, 1 To 17) As Variant

'1 Component
'2 Subgroup1
'3 Subgroup2
'4 circuit
'5 service tag
'6 point
'7 point_loc
'8 OD
'9 retire
'10 original date
'11 nominal
'12 subsequent date
'13 raw1
'14 raw2
'15 raw3
'16 avg
'17 a1 check
CancelTransfer = False
Z = 1
For X = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(X).Name <> "ListSheet" And Sheets(X).Name <> "Template" And Sheets(X).Name <> "BlankWS" And Sheets(X).Name <> "CalcSheet" And Sheets(X).Name <> "Homepage" And Sheets(X).Type <> 3 Then
        Y = 2
        Do
            If WorksheetFunction.IsNumber(Sheets(X).Cells(Y, 12)) Then
                
                XTArr(Z, 1) = Sheets(X).Cells(Y, 28)
                XTArr(Z, 2) = Sheets(X).Cells(Y, 22)
                XTArr(Z, 3) = ""
                XTArr(Z, 4) = Sheets(X).Cells(Y, 25)
                XTArr(Z, 5) = Sheets(X).Cells(Y, 26)
                XTArr(Z, 6) = Sheets(X).Cells(Y, 1)
                XTArr(Z, 7) = Sheets(X).Cells(Y, 2)
                XTArr(Z, 8) = Sheets(X).Cells(Y, 27)
                XTArr(Z, 9) = Sheets(X).Cells(Y, 3)
                XTArr(Z, 10) = Sheets(X).Cells(Y, 4)
                XTArr(Z, 11) = Sheets(X).Cells(Y, 5)
                XTArr(Z, 12) = Sheets(X).Cells(Y, 8)
                XTArr(Z, 13) = Sheets(X).Cells(Y, 9)
                XTArr(Z, 14) = Sheets(X).Cells(Y, 10)
                XTArr(Z, 15) = Sheets(X).Cells(Y, 11)
                XTArr(Z, 16) = Sheets(X).Cells(Y, 12)
                If Sheets(X).Cells(Y, 13) = "*" Then
                    XTArr(Z, 17) = "Fail"
                Else
                    XTArr(Z, 17) = "Pass"
                End If
                Z = Z + 1
            End If
            Y = Y + 1
        Loop Until Sheets(X).Cells(Y, 1) = ""
    End If
Next X
ChartName = Trim("Plot " & ActiveSheet.Range("A17"))
If Z > 0 Then
    WBWS.Show
    If CancelTransfer = False Then
        XTsh.Activate
        XTsh.Range(XTsh.Cells(3, 1), XTsh.Cells(Z + 2, 17)) = XTArr
    Else
        End
    End If
Else
    End
End If

End Sub

