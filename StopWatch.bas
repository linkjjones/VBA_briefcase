Attribute VB_Name = "StopWatch"
Option Explicit

Public Function StartTimer() As Double
    StartTimer = Timer
End Function

Public Function StopTimer(StartTime As Double) As String
    StopTimer = Format((Timer - StartTime) / 86400, "hh:mm:ss")
End Function
