Attribute VB_Name = "AccFunc"
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

Public Function SimpleListBoxItemSelected(ctlListBox As ListBox, _
                                          ReturnCol As Integer) As Long
    'check if a list item is selected
    If ctlListBox.ListIndex > -1 Then
        SimpleListBoxItemSelected = ctlListBox.Column(ReturnCol, ctlListBox.ListIndex)
    End If
    
End Function

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

Function WordAllowsAccessVBA(wd As Object) As Boolean
'Function WordAllowsAccessVBA(wd As Word.Document) As Boolean
    On Error Resume Next
    If Len(wd.VBProject.Name) Then
    End If
    
    If Err.Number Then
        WordAllowsAccessVBA = False
    Else
        WordAllowsAccessVBA = True
    End If
End Function

Public Sub StatusBar(ShowText As String)
    Dim varReturn As Variant
    If ShowText = "" Then
        ShowText = " "
    End If
    varReturn = SysCmd(acSysCmdSetStatus, ShowText)
End Sub

