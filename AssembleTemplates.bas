Attribute VB_Name = "AssembleTemplates"
Sub qry_units()

Dim qt As QueryTable
Dim tws As Worksheet
Dim ListSheetExists As Boolean


SQLConnect = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=I:\Technical Engineering Services\M&IS\Mining & Piping\ADMIN\DATA SHOP\Piping Integrity\PID_.accdb;"

SQLSelect = "Select Distinct tbl_Units.UnitNumber From tbl_Units Where tbl_Units.PriorityCircuits=True;"

For X = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(X).Name = "ListSheet" Then
        ListSheetExists = True
    End If
Next X
If ListSheetExists = False Then Sheets.Add().Name = "ListSheet"
Set tws = Sheets("ListSheet")
tws.Columns(1).ClearContents

Set qt = tws.QueryTables.Add(SQLConnect, tws.Range("A1"), SQLSelect)

With qt
    .BackgroundQuery = False
    .Refresh
End With


End Sub

Sub qry_CGs()

Dim qt As QueryTable
Dim tws As Worksheet


SQLConnect = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=I:\Technical Engineering Services\M&IS\Mining & Piping\ADMIN\DATA SHOP\Piping Integrity\PID_.accdb;"

SQLSelect = "Select tbl_CGs.CorrosionGroup From tbl_Units INNER JOIN tbl_CGs ON tbl_Units.UnitID = tbl_CGs.UnitID Where tbl_Units.UnitNumber = '" & Sheets("Homepage").Cells(3, 2) & "';"

Set tws = Sheets("ListSheet")
tws.Columns(2).ClearContents

Set qt = tws.QueryTables.Add(SQLConnect, tws.Range("B1"), SQLSelect)

With qt
    .BackgroundQuery = False
    .Refresh
End With

'disconnect database


End Sub

Sub qry_Circuits()

Dim qt As QueryTable
Dim tws As Worksheet


SQLConnect = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=I:\Technical Engineering Services\M&IS\Mining & Piping\ADMIN\DATA SHOP\Piping Integrity\PID_.accdb;"

SQLSelect = "Select tbl_Circuits.Circuit From (tbl_Units INNER JOIN tbl_CGs ON tbl_Units.UnitID = tbl_CGs.UnitID) INNER JOIN tbl_Circuits ON tbl_CGs.CGID = tbl_Circuits.CGID Where tbl_CGs.CorrosionGroup = '" & Sheets("Homepage").Cells(4, 2) & "';"

Set tws = Sheets("ListSheet")
tws.Columns(3).ClearContents

Set qt = tws.QueryTables.Add(SQLConnect, tws.Range("C1"), SQLSelect)

With qt
    .BackgroundQuery = False
    .Refresh
End With

End Sub


Sub qry_LineNumbers()

Dim qt As QueryTable
Dim tws As Worksheet


SQLConnect = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=I:\Technical Engineering Services\M&IS\Mining & Piping\ADMIN\DATA SHOP\Piping Integrity\PID_.accdb;"

SQLSelect = "Select tbl_Lines.LineNo From ((tbl_Units INNER JOIN tbl_CGs ON tbl_Units.UnitID = tbl_CGs.UnitID) INNER JOIN tbl_Circuits ON tbl_CGs.CGID = tbl_Circuits.CGID) INNER JOIN tbl_Lines ON tbl_Circuits.CircuitID = tbl_Lines.CircuitID Where tbl_CGs.CorrosionGroup = '" & Sheets("Homepage").Cells(4, 2) & "' AND tbl_Circuits.Circuit = '" & Sheets("Homepage").Cells(5, 2) & "';"

Set tws = Sheets("ListSheet")
tws.Columns(4).ClearContents

Set qt = tws.QueryTables.Add(SQLConnect, tws.Range("D1"), SQLSelect)

With qt
    .BackgroundQuery = False
    .Refresh
End With

End Sub


Sub qry_TMLs()

Dim qt As QueryTable
Dim tws As Worksheet
Dim ButtonRng As Range

SQLConnect = "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=I:\Technical Engineering Services\M&IS\Mining & Piping\ADMIN\DATA SHOP\Piping Integrity\PID_.accdb;"

SQLSelect = "Select tbl_TMLs.TML, tbl_TMLs.TMLLocation, tbl_TMLs.RetirementLimit, tbl_TMLs.OriginalDate, tbl_TMLs.OriginalThickness, tbl_Circuits.InspectionEffectiveness, tbl_TMLs.OD  From (((tbl_Units INNER JOIN tbl_CGs ON tbl_Units.UnitID = tbl_CGs.UnitID) INNER JOIN tbl_Circuits ON tbl_CGs.CGID = tbl_Circuits.CGID) INNER JOIN tbl_Lines ON tbl_Circuits.CircuitID = tbl_Lines.CircuitID) INNER JOIN tbl_TMLs ON tbl_Lines.LineID = tbl_TMLs.LineID " & _
            "Where tbl_CGs.CorrosionGroup = '" & Sheets("Homepage").Cells(4, 2) & "' AND tbl_Circuits.Circuit = '" & Sheets("Homepage").Cells(5, 2) & "' AND tbl_Lines.LineNo = '" & Sheets("Homepage").Cells(6, 2) & "' AND tbl_TMLs.[TML Type] <> 'IDM - Discontinue Monitoring' AND tbl_TMLs.[TML Type] <> 'IDM - Delete';"

Set tws = Sheets("Homepage")
tws.Columns("D:M").Delete Shift:=xlToLeft

Set qt = tws.QueryTables.Add(SQLConnect, tws.Range("D8"), SQLSelect)

With qt
    .BackgroundQuery = False
    .Refresh
End With

Call FormatHomepage(tws.UsedRange.Rows.Count)
'tws.Range("K8") = "Select": tws.Range("K8").Font.Bold = True
'tws.Range("L8") = "G/NG Compliant": tws.Range("L8").Font.Bold = True
'tws.Range("M8") = "Component Type": tws.Range("M8").Font.Bold = True

'With tws.Range("M9:M" & tws.UsedRange.Rows.Count).Validation
    'If tws.Range("B2") = "UT" Then
        '.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Pipe, Elbow, Reducer 1, Reducer 2, T1, T2, T3, T4"
    'ElseIf tws.Range("B2") = "RT" Then
        '.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Pipe, Elbow, Tee (Fitting), Tee (Stubin), Reducer, Pipe+Flange, SBC"
    'ElseIf tws.Range("B2") = "UT (Mix Point)" Then
        '.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Pipe, Elbow, Reducer 1, Reducer 2, T1, T2, T3, T4, Mixing Tee"
    'ElseIf tws.Range("B2") = "UT (Injection Point)" Then
        '.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Pipe, Elbow, Reducer 1, Reducer 2, T1, T2, T3, T4, Tee w/ Quill"
    'ElseIf tws.Range("B2") = "UT (Stagnant Zone)" Or tws.Range("B2") = "UT (Deadleg)" Then
        '.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Horizontal Pipe, Vertical Down Pipe, Vertical Up Pipe, Pipe End, Elbow, Reducer 1, Reducer 2, T1, T2, T3, T4, Horizontal Tee, Downward Tee, Upward Tee"
    'End If
        
'End With

'tws.Range("K9:M" & tws.UsedRange.Rows.Count).Style = "Input"
'With Columns("D:M").EntireColumn
    '.AutoFit
    '.HorizontalAlignment = xlCenter
'End With
'tws.Columns("L:L").Hidden = True

'Set ButtonRng = ActiveSheet.Range("D" & ActiveSheet.UsedRange.Rows.Count + 1 & ":M" & ActiveSheet.UsedRange.Rows.Count + 2)
'ActiveSheet.Buttons.Delete
'On Error Resume Next
'ActiveSheet.Buttons("Btn1").Delete
'On Error GoTo 0
'ActiveSheet.Buttons.Add(ButtonRng.Left, ButtonRng.Top, ButtonRng.Width, ButtonRng.Height).Select
'With Selection
  '.OnAction = "AssembleTemplates"
  '.Caption = "Assemble Template"
  '.Name = "Btn1"
'End With
'Range("B6").Select

End Sub

Sub AssembleTemplates()

'1. loop through TML list and input information into
    'a. array for GIP compliant TMLs
    'b. array for Non-GIP compliant TMLs
'2. insert a worksheet with the line number as the title
'3. add templates from hidden worksheets according to component type and inspection method (add G first then NG)
'4. fill templates with information from arrays
'5. query last subsequent date/thickness for each TML (if it exists)

Dim Garr() As Variant 'GIP Array
Dim Gcnt As Integer
Dim NGcnt As Integer
Dim NGarr() As Variant 'Non-GIP Array
Dim t As Integer 'tml
Dim p As Integer 'point
Dim aSh As Worksheet
Dim tSh As Worksheet
Dim dsh As Worksheet
Dim dShRows As Long
Dim dr As Variant 'drawing reference
Dim refDWG As Object
Dim dwgrange As Range
Dim drGArr() As Variant
Dim drNGArr() As Variant
Dim hideButton As OLEObject


'Application.ScreenUpdating = False

Set aSh = ActiveSheet
Set tSh = Sheets("Template")


If aSh.Range("B1") = "" Or aSh.Range("B2") = "" Or aSh.Range("B3") = "" Or aSh.Range("B4") = "" Or aSh.Range("B5") = "" Or aSh.Range("B6") = "" Then
    MsgBox "You must fill in all necessary information.", vbCritical
    End
End If

'****************************************************Section 1************************************************
'count Gip/Non Gip TMLs before dimensioning arrays

'make all TML NG for now....
t = 1
Do Until aSh.Cells(t + 8, 4) = ""
    aSh.Cells(t + 8, 12) = "NG"
    t = t + 1
Loop

t = 1
Do Until aSh.Cells(t + 8, 4) = ""
    If aSh.Cells(t + 8, 11) = "*" And aSh.Cells(t + 8, 12) = "G" Then
        Gcnt = Gcnt + 1
    ElseIf aSh.Cells(t + 8, 11) = "*" And aSh.Cells(t + 8, 12) = "NG" Then
        NGcnt = NGcnt + 1
    End If
    t = t + 1
Loop
If Gcnt > 0 Then
    ReDim Garr(1 To Gcnt, 1 To 8)
    ReDim drGArr(1 To Gcnt)
End If
If NGcnt > 0 Then
    ReDim NGarr(1 To NGcnt, 1 To 8)
    ReDim drNGArr(1 To NGcnt)
End If

If NGcnt = 0 Then
    MsgBox "You have not selected any TMLs to produce a template with.", vbCritical
    End
End If

'populate arrays
Gcnt = 1
NGcnt = 1
t = 1
Do Until aSh.Cells(t + 8, 4) = ""
    If aSh.Cells(t + 8, 11) = "*" And aSh.Cells(t + 8, 12) = "G" Then
        Garr(Gcnt, 1) = aSh.Cells(t + 8, 4) 'TML
        Garr(Gcnt, 2) = aSh.Cells(t + 8, 5) 'TMLLocation
        Garr(Gcnt, 3) = aSh.Cells(t + 8, 6) 'RetirementLimit
        Garr(Gcnt, 4) = aSh.Cells(t + 8, 7) 'OriginalDate
        Garr(Gcnt, 5) = aSh.Cells(t + 8, 8) 'OriginalThickness
        Garr(Gcnt, 6) = aSh.Cells(t + 8, 9) 'InspectionEffectiveness
        Garr(Gcnt, 7) = aSh.Cells(t + 8, 10) 'OD
        Garr(Gcnt, 8) = aSh.Cells(t + 8, 13) 'ComponentType
        If Garr(Gcnt, 4) >= aSh.Range("B1") Then
            MsgBox "Your inspection date occur before or be the same as your original date.", vbCritical
            End
        End If
        Gcnt = Gcnt + 1
    ElseIf aSh.Cells(t + 8, 11) = "*" And aSh.Cells(t + 8, 12) = "NG" Then
        NGarr(NGcnt, 1) = aSh.Cells(t + 8, 4) 'TML
        NGarr(NGcnt, 2) = aSh.Cells(t + 8, 5) 'TMLLocation
        NGarr(NGcnt, 3) = aSh.Cells(t + 8, 6) 'RetirementLimit
        NGarr(NGcnt, 4) = aSh.Cells(t + 8, 7) 'OriginalDate
        NGarr(NGcnt, 5) = aSh.Cells(t + 8, 8) 'OriginalThickness
        NGarr(NGcnt, 6) = aSh.Cells(t + 8, 9) 'InspectionEffectiveness
        NGarr(NGcnt, 7) = aSh.Cells(t + 8, 10) 'OD
        NGarr(NGcnt, 8) = aSh.Cells(t + 8, 13) 'ComponentType
        If NGarr(NGcnt, 1) = "" Or NGarr(NGcnt, 2) = "" Or NGarr(NGcnt, 3) = "" Or NGarr(NGcnt, 4) = "" Or NGarr(NGcnt, 5) = "" Or NGarr(NGcnt, 6) = "" Or NGarr(NGcnt, 7) = "" Then
            MsgBox "There is missing TML data in one of your selections, contact a Syncrude Piping Integrity Team member for assistance.", vbCritical
            End
        ElseIf NGarr(NGcnt, 8) = "" Then
            MsgBox "You must provide the Component Type for TML " & NGarr(NGcnt, 1), vbCritical
            End
        End If
        If NGarr(NGcnt, 4) >= aSh.Range("B1") Then
            MsgBox "Your inspection date cannot occur before or be the same as your original date.", vbCritical
            End
        End If
        NGcnt = NGcnt + 1
    End If
    t = t + 1
Loop

'****************************************************Section 2************************************************
For X = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(X).Name = Trim(aSh.Range("B6")) & " " & Left(aSh.Range("B2"), 2) Then
        TempSheetExists = True
    End If
Next X
If TempSheetExists = False Then
    'Sheets.Add().Name = Trim(aSh.Range("B6"))
    
    'Call AddCodeToTemplate
     Sheets("BlankWS").Visible = True
     Sheets("BlankWS").Copy After:=Sheets(Sheets.Count)
     Sheets("BlankWS").Visible = False
     Sheets(Sheets.Count).Name = Trim(aSh.Range("B6")) & " " & Left(aSh.Range("B2"), 2)
Else
    MsgBox "You can only have one template per line number. Delete the first one and try again.", vbCritical
    End
End If
Set dsh = Sheets(Trim(aSh.Range("B6")) & " " & Left(aSh.Range("B2"), 2))
TempSheetExists = False

'****************************************************Section 3************************************************
'develop instruction string (i.e. InspectionMethod-ComponentType-Size(1or2)/#ofPlanes-Localized/Uniform-Std/Med/H

'add NG TMLs first (i.e. need to use templates)
For X = 1 To UBound(NGarr)
    im = aSh.Range("B2")
    If NGarr(X, 8) = "Pipe" Then
        ct = "P"
    ElseIf NGarr(X, 8) = "Elbow" Then
        ct = "E"
    ElseIf NGarr(X, 8) = "Tee (Fitting)" Then
        ct = "T"
    ElseIf NGarr(X, 8) = "Tee (Stubin)" Then
        ct = "SI"
    ElseIf NGarr(X, 8) = "Reducer 1" Then
        ct = "R1"
    ElseIf NGarr(X, 8) = "Reducer 2" Then
        ct = "R2"
    ElseIf NGarr(X, 8) = "Reducer" Then
        ct = "R"
    ElseIf NGarr(X, 8) = "Pipe+Flange" Then
        ct = "F"
    ElseIf NGarr(X, 8) = "Mixing Tee" Then
        ct = "TM"
    ElseIf NGarr(X, 8) = "Tee w/ Quill" Then
        ct = "TQ"
    ElseIf NGarr(X, 8) = "Horizontal Pipe" Then
        ct = "PH"
    ElseIf NGarr(X, 8) = "Vertical Down Pipe" Then
        ct = "PD"
    ElseIf NGarr(X, 8) = "Vertical Up Pipe" Then
        ct = "PU"
    ElseIf NGarr(X, 8) = "Pipe End" Then
        ct = "PE"
    ElseIf NGarr(X, 8) = "Downward Tee" Then
        ct = "TD"
    ElseIf NGarr(X, 8) = "Upward Tee" Then
        ct = "TU"
    ElseIf NGarr(X, 8) = "Horizontal Tee" Then
        ct = "TH"
    Else
        ct = NGarr(X, 8)
    End If
    
    If Left(im, 2) = "UT" Then
        If NGarr(X, 7) < 10 Then s = 1
        If NGarr(X, 7) >= 10 Then s = 2
    ElseIf im = "RT" Then
        s = InputBox("How many planes did you shoot TML #" & Trim(NGarr(X, 1)) & " in?", "Number of Planes for RT")
    End If
    
    If im = "UT" Then
        cml = InStr(1, UCase(NGarr(X, 6)), "LOCALIZED"): If cml <> 0 Then cm = "L"
        cmu = InStr(1, UCase(NGarr(X, 6)), "UNIFORM"): If cmu <> 0 Then cm = "U"
    ElseIf im = "UT (Mix Point)" Then
        If ct = "TM" Then
            cm = "MP"
        Else
            cml = InStr(1, UCase(NGarr(X, 6)), "LOCALIZED"): If cml <> 0 Then cm = "L"
            cmu = InStr(1, UCase(NGarr(X, 6)), "UNIFORM"): If cmu <> 0 Then cm = "U"
        End If
    ElseIf im = "UT (Injection Point)" Then
        If ct = "TQ" Then
            cm = "IP"
        Else
            cml = InStr(1, UCase(NGarr(X, 6)), "LOCALIZED"): If cml <> 0 Then cm = "L"
            cmu = InStr(1, UCase(NGarr(X, 6)), "UNIFORM"): If cmu <> 0 Then cm = "U"
        End If
    ElseIf im = "UT (Stagnant Zone)" Or im = "UT (Deadleg)" Then
        If ct = "PH" Or ct = "PD" Or ct = "PU" Or ct = "PE" Or ct = "TH" Or ct = "TD" Or ct = "TU" Then
            cm = "SZ"
        Else
            cml = InStr(1, UCase(NGarr(X, 6)), "LOCALIZED"): If cml <> 0 Then cm = "L"
            cmu = InStr(1, UCase(NGarr(X, 6)), "UNIFORM"): If cmu <> 0 Then cm = "U"
        End If
    End If
    ils = InStr(1, NGarr(X, 6), "Standard"): If ils <> 0 Then il = "S"
    ilm = InStr(1, NGarr(X, 6), "Medium"): If ilm <> 0 Then il = "M"
    ilh = InStr(1, NGarr(X, 6), "High"): If ilh <> 0 Then il = "H"
    
    If Left(im, 2) = "UT" Then
        'If (Left(aSh.Range("B5"), 2) = "MP" Or Left(aSh.Range("B5"), 2) = "IP") And NGarr(X, 8) = "T3" Then
            'templatestring = im & "-" & ct & "-" & s & "-" & "MP"
        'Else
        templatestring = Left(im, 2) & "-" & ct & "-" & s & "-" & cm & "-" & il
        'End If
    ElseIf im = "RT" Then
        templatestring = im & "-" & ct & "-" & s
    End If
    
    dShRows = dsh.UsedRange.Rows.Count
    If dShRows = 1 Then
        tSh.Range("A1:U1").Copy dsh.Cells(1, 1)
        'dsh.Paste
        dsh.Rows(1).RowHeight = 60
        dShRows = dShRows + 1
 
        Set hideButton = ActiveSheet.OLEObjects.Add(ClassType:="Forms.CheckBox.1", Link:=False, DisplayAsIcon:=False, Left:=5, Top:=9.75, Width:=111, Height:=22.5)
        With hideButton.Object
            .Caption = "Hide Details"
            .Font.Size = 14
            .Font.Italic = True
            .Font.Bold = True
            .ForeColor = 255
        End With
        
    End If
        
    tSh.Range("A2:U71").Copy dsh.Cells(dShRows, 1)
        
    For y = dShRows To dsh.UsedRange.Rows.Count - 1
        If dsh.Cells(y, 1) <> "min" Then
            dsh.Cells(y, 1) = Trim(NGarr(X, 1)) & dsh.Cells(y, 1)
        ElseIf dsh.Cells(y, 1) = "min" And Left(Sheets("Homepage").Range("B5"), 2) = "SB" Then
            dsh.Cells(y, 1) = Trim(NGarr(X, 1))
        End If
        dsh.Cells(y, 2) = NGarr(X, 2) & dsh.Cells(y, 2)
        dsh.Cells(y, 3) = NGarr(X, 3) & dsh.Cells(y, 3)
        dsh.Cells(y, 4) = NGarr(X, 4) & dsh.Cells(y, 4)
        dsh.Cells(y, 5) = NGarr(X, 5) & dsh.Cells(y, 5)
        dsh.Cells(y, 8) = aSh.Range("B1")
        dsh.Cells(y, 23) = aSh.Range("B2")
        dsh.Cells(y, 24) = aSh.Range("B3")
        dsh.Cells(y, 25) = aSh.Range("B4")
        dsh.Cells(y, 26) = aSh.Range("B5")
        dsh.Cells(y, 27) = aSh.Range("B6")
        dsh.Cells(y, 28) = NGarr(X, 7)
        dsh.Cells(y, 29) = NGarr(X, 8)
    Next y
    
    For y = 22 To 183
        If tSh.Cells(1, y) = templatestring Then Exit For
    Next y
    
    dr = "D" & tSh.Cells(72, y)
    drNGArr(X) = dr
    tSh.Range(tSh.Cells(2, y), tSh.Cells(71, y)).Copy dsh.Cells(dShRows, 22)
    
Next X

'add G TMLs second (i.e. don't need imbedded templates)

dShRows = dsh.UsedRange.Rows.Count

For y = dShRows - 1 To 2 Step -1
    If dsh.Cells(y, 22) = "Y" Or dsh.Cells(y, 22) = "" Then
        Rows(y).Delete
    End If
Next y

nextcell = 2
For y = 1 To UBound(drNGArr)
    
    mergedrange = dsh.Cells(nextcell, 21).MergeArea.Address
    nextcellarr = Split(mergedrange, "$", -1)
    nextcell = nextcellarr(UBound(nextcellarr)) + 1
    Set dwgrange = dsh.Range(mergedrange)
    
    Set refDWG = tSh.Pictures(drNGArr(y))
    refDWG.Copy
    dsh.Paste
    
    dsh.Columns(19).ColumnWidth = 55
    dsh.Columns(20).ColumnWidth = 55
    dsh.Columns(21).ColumnWidth = 55
    
    With dsh.Pictures(dsh.Pictures.Count)
        .Height = .Height / 1.5
        .Width = .Width / 1.5
        
        If dwgrange.Height < .Height Then
            RowHeightReq = .Height / dwgrange.Rows.Count
            For Each r In dwgrange.Rows
                r.RowHeight = RowHeightReq + (RowHeightReq / 2)
            Next
        End If
        .Top = (dwgrange.Top + (dwgrange.Height / 2)) - (.Height / 2)
        .Left = dwgrange.Left + ((dwgrange.Width - .Width) / 2)
        .OnAction = "ImageZoom"
    End With
    
Next y

'If dsh.Cells(dsh.UsedRange.Rows.Count, 1) <> "" Then
    'AddRow = 1
'End If

'With dsh.Range(dsh.Cells(dsh.UsedRange.Rows.Count + AddRow, 1), dsh.Cells(dsh.UsedRange.Rows.Count + AddRow, 21)).Borders(xlEdgeTop)
    '.LineStyle = xlContinuous
    '.ColorIndex = xlAutomatic
    '.TintAndShade = 0
    '.Weight = xlThick
'End With
For y = 2 To dsh.UsedRange.Rows.Count + 1
    If dsh.Range("A" & y) <> "min" Then
        If Left(Right(dsh.Range("A" & y), 2), 1) <> Left(Right(dsh.Range("A" & y - 1), 2), 1) Then
            With dsh.Range("A" & y & ":" & "U" & y).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlThick
            End With
        End If
    End If
Next y

dsh.Columns("A:R").EntireColumn.AutoFit
dsh.Columns("V").Delete
dsh.Columns("V:AC").Hidden = True
dsh.Cells(1, 22) = "Inspection Method": dsh.Cells(1, 23) = "Unit": dsh.Cells(1, 24) = "Corrosion Group": dsh.Cells(1, 25) = "Circuit": dsh.Cells(1, 26) = "Line Number": dsh.Cells(1, 27) = "OD": dsh.Cells(1, 28) = "Component Type"
If im = "RT" Then dsh.Columns("J:K").Hidden = True
If Left(im, 2) = "UT" Then dsh.Columns("T:T").Hidden = True
With dsh.Columns("R:R")
    .ColumnWidth = 60
    .WrapText = True
End With
'worksheet password Dh1986
dsh.Range("I:I,J:J,K:K,R:R,S:S,T:T").Locked = False
dsh.Protect "Dh1986", False, , , True
If Sheets("Homepage").Shapes("IDMConnected").ControlFormat.Value = 1 Then
    Call qry_LastInspectionInfo
    Call AddLastestInpsectionInfoIntoTemplate
End If

ActiveWindow.DisplayGridlines = False

End Sub

Sub qry_LastInspectionInfo()

Dim qt As QueryTable
Dim tws As Worksheet
Dim TempSheetExists As Boolean
Dim hSh As Worksheet

Set hSh = Sheets("Homepage")

SQLConnect = "ODBC;driver={SQL Server};server=SYNSQL05;UID=;PWD=;DSN=;Trusted_Connection=Yes;Database={IDMS};AnsiNpw=No;QuotedId=No;Regional=Yes "

SQLSelect = "SELECT g.point, g.MaxDate, point_subs_thick.rawavg " & _
            "FROM (SELECT point, MAX(subs_date) AS MaxDate " & _
                  "FROM point_subs_thick " & _
                  "WHERE service_tag = '" & hSh.Range("B6") & "' " & _
                  "GROUP BY point) AS g " & _
            "JOIN point_subs_thick ON g.point = point_subs_thick.point AND g.MaxDate = point_subs_thick.subs_date " & _
            "WHERE service_tag = '" & hSh.Range("B6") & "' " & _
            "ORDER BY g.point;"

For X = 1 To ActiveWorkbook.Sheets.Count
    If Sheets(X).Name = "TempSheet" Then
        TempSheetExists = True
    End If
Next X
If TempSheetExists = False Then Sheets.Add().Name = "TempSheet"
Set tws = Sheets("TempSheet")

Set qt = tws.QueryTables.Add(SQLConnect, tws.Range("A1"), SQLSelect)

With qt
    .BackgroundQuery = False
    .Refresh
End With


End Sub

Sub AddLastestInpsectionInfoIntoTemplate()

Dim tws As Worksheet
Dim dws As Worksheet

Application.DisplayAlerts = False

Set tws = Sheets("TempSheet")
Set dws = Sheets(Trim(Sheets("Homepage").Range("B6")) & " " & Left(Sheets("Homepage").Range("B2"), 2))


For X = 2 To tws.UsedRange.Rows.Count
    For y = 2 To dws.UsedRange.Rows.Count
        If tws.Cells(X, 1) = dws.Cells(y, 1) Then
            dws.Cells(y, 6) = tws.Cells(X, 2)
            dws.Cells(y, 7) = tws.Cells(X, 3)
        End If
    Next y
Next X

tws.Delete

End Sub

Sub AddCodeToTemplate()

Dim aSh As Worksheet

newHour = Hour(Now())
newMinute = Minute(Now())
newSecond = Second(Now()) + 1
waitTime = TimeSerial(newHour, newMinute, newSecond)
Application.Wait waitTime

Set aSh = Sheets("HomePage")

    With ActiveWorkbook.VBProject.VBComponents(Sheets(Trim(aSh.Range("B6"))).CodeName).CodeModule
            .InsertLines 1, "Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)"
            .InsertLines 2, "If (target.Column = 19 Or target.Column = 20) And Cells(target.Row, 1) <> """" Then"
            .InsertLines 3, "Call InsertTMLPhoto"
            .InsertLines 4, "Cancel = True"
            .InsertLines 5, "ElseIf target.Column = 1 And target.Interior.TintAndShade = -0.249977111117893 Then"
            .InsertLines 6, "TML = Left(Cells(target.Row - 1, 1), Len(Cells(target.Row - 1, 1)) - 2)"
            .InsertLines 7, "Call InsertMinTML"
            .InsertLines 8, "Cancel = True"
            .InsertLines 9, "ElseIf target.Column = 1 And Left(Sheets(""Homepage"").Range(""B5""), 2) = ""SB"" Then"
            .InsertLines 10, "Call InsertMinTML"
            .InsertLines 11, "Cancel = True"
            .InsertLines 12, "ElseIf target.Column = 14 And target.Row > 1 Then"
            .InsertLines 13, "Call BandCharts(target)"
            .InsertLines 14, "Cancel = True"
            .InsertLines 15, "ElseIf target.Column = 15 And target.Row > 1 Then"
            .InsertLines 16, "Call GuessPipeSchedule(target)"
            .InsertLines 17, "Cancel = True"
            .InsertLines 18, "End If"
            .InsertLines 19, "End Sub"
            .InsertLines 20, "Private Sub CheckBox1_Click()"
            .InsertLines 21, "If CheckBox1 = True Then"
            .InsertLines 22, "Columns(""B:H"").Hidden = True"
            .InsertLines 23, "Else"
            .InsertLines 24, "Columns(""B:H"").Hidden = False"
            .InsertLines 25, "End If"
            .InsertLines 26, "End Sub"

    End With

End Sub

Sub FormatHomepage(nRows)

Dim tws As Worksheet
Dim ButtonRng As Range


Set tws = Sheets("Homepage")
'tws.Columns("D:M").Delete Shift:=xlToLeft

tws.Range("D8") = "TML": tws.Range("D8").Font.Bold = True
tws.Range("E8") = "TMLLocation": tws.Range("E8").Font.Bold = True
tws.Range("F8") = "RetirementLimit": tws.Range("F8").Font.Bold = True
tws.Range("G8") = "OriginalDate": tws.Range("G8").Font.Bold = True
tws.Range("H8") = "OriginalThickness": tws.Range("H8").Font.Bold = True
tws.Range("I8") = "InspectionEffectiveness": tws.Range("I8").Font.Bold = True
tws.Range("J8") = "OD": tws.Range("J8").Font.Bold = True
tws.Range("K8") = "Select": tws.Range("K8").Font.Bold = True
tws.Range("L8") = "G/NG Compliant": tws.Range("L8").Font.Bold = True
tws.Range("M8") = "Component Type": tws.Range("M8").Font.Bold = True

With tws.Range("M9:M" & nRows).Validation
    If tws.Range("B2") = "UT" Then
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Pipe, Elbow, Reducer 1, Reducer 2, T1, T2, T3, T4"
    ElseIf tws.Range("B2") = "RT" Then
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Pipe, Elbow, Tee (Fitting), Tee (Stubin), Reducer, Pipe+Flange, SBC"
    ElseIf tws.Range("B2") = "UT (Mix Point)" Then
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Pipe, Elbow, Reducer 1, Reducer 2, T1, T2, T3, T4, Mixing Tee"
    ElseIf tws.Range("B2") = "UT (Injection Point)" Then
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Pipe, Elbow, Reducer 1, Reducer 2, T1, T2, T3, T4, Tee w/ Quill"
    ElseIf tws.Range("B2") = "UT (Stagnant Zone)" Or tws.Range("B2") = "UT (Deadleg)" Then
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Horizontal Pipe, Vertical Down Pipe, Vertical Up Pipe, Pipe End, Elbow, Reducer 1, Reducer 2, T1, T2, T3, T4, Horizontal Tee, Downward Tee, Upward Tee"
    End If
        
End With

tws.Range("K9:M" & nRows).Style = "Input"
With Columns("D:M").EntireColumn
    .AutoFit
    .HorizontalAlignment = xlCenter
End With
tws.Columns("L:L").Hidden = True

Set ButtonRng = ActiveSheet.Range("D" & nRows + 1 & ":M" & nRows + 2)
'ActiveSheet.Buttons.Delete
On Error Resume Next
ActiveSheet.Buttons("Btn1").Delete
On Error GoTo 0
ActiveSheet.Buttons.Add(ButtonRng.Left, ButtonRng.Top, ButtonRng.Width, ButtonRng.Height).Select
With Selection
  .OnAction = "AssembleTemplates"
  .Caption = "Assemble Template"
  .Name = "Btn1"
End With
Range("B6").Select

End Sub

Sub HomepageListValidations()

Dim tws As Worksheet
On Error Resume Next

Set tws = Sheets("Homepage")

If tws.Shapes("PIDConnected").ControlFormat.Value = 1 Then
    With tws.Range("B3").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Units"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With tws.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=CG"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With tws.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=Circuits"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With tws.Range("B6").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=LineNumbers"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

Else
    For X = 3 To 6
        With tws.Range("B" & X).Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
            :=xlBetween
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Next X
End If


End Sub
