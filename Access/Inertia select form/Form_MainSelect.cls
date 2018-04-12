VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MainSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Box0_Click()
    'set colors tob buttons
'    Dim ctl As Control
'    Dim color As Long
'    Debug.Print Box0.BackColor
'    color = Box0.BackColor
'    For Each ctl In Me
'        If InStr(1, ctl.Name, 1) > 0 Then
'            ctl.BackColor = color
'        End If
'    Next ctl
    
    'Get distance between lnTop/lnBtm
'    Debug.Print lnTop.Top - lnBtm.Top
End Sub

Private Sub frmTabSelect_Click()
    Call AttractBar
    TabCtl.Value = IIf(frmTabSelect.Value < 2, frmTabSelect.Value, TabCtl.Value)
End Sub

Private Sub AttractBar()
    Dim curr            As Long
    Dim dest            As Long
    Dim origin          As Long
    Dim totalDistance   As Long
    Dim distanceTraveled As Long
    Dim direction       As Integer
    Dim step            As Long
    Dim speed           As Integer
    Dim percentTraveled As Double
    
    Select Case frmTabSelect
        Case Is = 1
            dest = togInspection.Left
        Case Is = 2
            dest = togIntervention.Left
        Case Is = 3
            dest = togOption.Left
    End Select
    
    origin = bar.Left

    If origin > dest Then
        'move left
        direction = -1
    ElseIf origin < dest Then
        'move right
        direction = 1
    Else
        'stay
        Exit Sub
    End If
    
    totalDistance = Abs(dest - origin)
    curr = origin
    
    'uses inertia until within 20 ticks if destination
    Do While percentTraveled * 100 < 100
        'get percent traveled
        distanceTraveled = Abs(curr - origin)
'        Debug.Print distanceTraveled & " distanceTraveled"
        percentTraveled = Round(distanceTraveled / totalDistance, 2)
'        Debug.Print percentTraveled * 100 & " percentTraveled"
        step = Parabolic.Inertia(percentTraveled)
'        Debug.Print step & " step"
        step = step * direction
'        Debug.Print step & " step + direction"
        curr = curr + step
'        Debug.Print curr & " new position"
'        Debug.Print " "
        bar.Left = curr
        Me.Repaint
    Loop
    bar.Left = dest
    
    
End Sub

Private Sub frmComponentGrp_Click()
    Call AttractBars
End Sub

Private Sub AttractBars()
    Dim curr            As Long
    Dim dest            As Long
    Dim origin          As Long
    Dim totalDistance   As Long
    Dim distanceTraveled As Long
    Dim direction       As Integer
    Dim step            As Long
    Dim speed           As Integer
    Dim percentTraveled As Double
    Dim btmLnOffset     As Long
    Dim topLnOffset     As Long
    
    topLnOffset = -14
    btmLnOffset = 490
    
    Select Case frmComponentGrp
        Case Is = 1
            dest = togOuterUC1.Top
        Case Is = 2
            dest = togEqBar1.Top
        Case Is = 3
            dest = togSwingFrame1.Top
        Case Is = 4
            dest = togIceLugs1.Top
    End Select
    
    origin = lnTop.Top

    If origin > dest Then
        'move left
        direction = -1
    ElseIf origin < dest Then
        'move right
        direction = 1
    Else
        'stay
        Exit Sub
    End If
    
    totalDistance = Abs(dest - origin)
    curr = origin
    
    'uses inertia until within 20 ticks if destination
    Do While percentTraveled * 100 < 100
        'get percent traveled
        distanceTraveled = Abs(curr - origin)
'        Debug.Print distanceTraveled & " distanceTraveled"
        percentTraveled = Round(distanceTraveled / totalDistance, 2)
'        Debug.Print percentTraveled * 100 & " percentTraveled"
        step = Parabolic.Inertia(percentTraveled)
'        Debug.Print step & " step"
        step = step * direction
'        Debug.Print step & " step + direction"
        curr = curr + step
'        Debug.Print curr & " new position"
'        Debug.Print " "
        lnTop.Top = curr + topLnOffset
        lnBtm.Top = curr + btmLnOffset
        Me.Repaint
    Loop
    lnTop.Top = dest + topLnOffset
    lnBtm.Top = dest + btmLnOffset
    
End Sub
