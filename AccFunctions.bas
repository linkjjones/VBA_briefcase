Attribute VB_Name = "AccFunctions"
Option Compare Database
Option Explicit

Global g_Enabled As Boolean
Global boolShift As Boolean

'Private Sub chkSelect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    'detect if the shift key is down (set boolShift)
'    boolShift = -Shift
'End Sub
'
'If boolShift then...

'Hide Nav pane
Public Function HideNavPane() As Byte
    DoCmd.SelectObject acTable, "tmpLineConf", True
    DoCmd.RunCommand acCmdWindowHide
End Function

'Show Nav pane
Public Sub UnHideNavPane()
    DoCmd.SelectObject acTable, "tmpLineConf", True
End Sub

'Minimize nav pane
Public Sub MinimizeNavPane()
    DoCmd.SelectObject acTable, , True
    DoCmd.Minimize
End Sub

'Clear all controls on a given form
Public Sub ClearAllControls(frm As Form, Optional Tag As String)
    Dim ctl As Control
    
    If Not Tag = "" Then
        For Each ctl In frm
            If ctl.Tag = Tag Then
                If Left(ctl.Name, 3) = "txt" Then
                    ctl = Null
                ElseIf Left(ctl.Name, 3) = "cbo" Then
                    ctl = Null
                ElseIf Left(ctl.Name, 3) = "chk" Then
                    ctl = False
                ElseIf Left(ctl.Name, 3) = "lst" Then
                    ctl.RowSource = ""
                ElseIf Left(ctl.Name, 3) = "tog" Then
                    ctl = 0
                End If
            End If
        Next ctl
    Else
        For Each ctl In frm
            If Left(ctl.Name, 3) = "txt" Then
                ctl = Null
            ElseIf Left(ctl.Name, 3) = "cbo" Then
                    ctl = Null
            ElseIf Left(ctl.Name, 3) = "lst" Then
                    ctl.RowSource = ""
            ElseIf Left(ctl.Name, 3) = "chk" Then
                    ctl = False
            ElseIf Left(ctl.Name, 3) = "tog" Then
                    ctl = 0
            End If
        Next ctl
    End If
    
End Sub

Public Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim LB As Long
    Dim UB As Long

    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I
        ' cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occasions, LBound is 0 and
        ' UBound is -1.
        ' To accommodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(Arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function

Public Function SimpleListBoxItemSelected(ctlListBox As ListBox, _
                                          ReturnCol As Integer) As Long
    'check if a list item is selected
    If ctlListBox.ListIndex > -1 Then
        SimpleListBoxItemSelected = ctlListBox.Column(ReturnCol, ctlListBox.ListIndex)
    End If
    
End Function

Public Sub SayThis(Sentence As String)
    Dim s As Object
    Dim vol As Long
    
    Set s = CreateObject("SAPI.SpVoice")
    
    'Get current volume
    vol = s.volume
    'set higher volume
    s.volume = 100
    'Say it
    s.Speak Sentence
    'Set volume to original level
    s.volume = vol
    'Cleanup
    Set s = Nothing
    
End Sub

' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish


Sub SleepFor(lngMilliSec As Long)
    If lngMilliSec > 0 Then
        Call sapiSleep(lngMilliSec)
    End If
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'        MsgBox "Please use the Logout button to close the application!!!!", vbInformation, ""
'        DoCmd.CancelEvent
'End Sub

Function WordAllowsAccessVBA(wd As Word.Document) as boolean
    On Error Resume Next
    If Len(wd.VBProject.Name) Then
    End If
    
    If Err.Number Then
        WordAllowsAccessVBA = False
    Else
        WordAllowsAccessVBA = True
    End If
End Function
