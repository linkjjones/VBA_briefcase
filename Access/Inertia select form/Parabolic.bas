Attribute VB_Name = "Parabolic"
Option Compare Database
Option Explicit

Public Function Inertia(PercentScrolled As Double) As Long
    Dim k As Integer
    Dim h As Integer
    Dim a As Integer
    'returns: big, small, big
    
    a = 20 'rate of change
    h = 0.5 'negative y: constant (0 to 1: 50%=0.5)
    k = 100  'x fastest (midpoint)
    
    Inertia = Round(a * (PercentScrolled + h) ^ 2 + k)
'    Debug.Print Inertia
End Function

