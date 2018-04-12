VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataSelect 
   Caption         =   "Select Data"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6600
   OleObjectBlob   =   "DataSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAll_Click()
    Call GetMSLData.GetSpools
    Call GetSpoolHistory
    Unload Me
End Sub

Private Sub btnAllAN_Click()
    Call GetMSLData.GetSpools(1)
    Call GetSpoolHistory(1)
    Unload Me
End Sub

Private Sub btnAllML_Click()
    Call GetMSLData.GetSpools(2)
    Call GetSpoolHistory(2)
    Unload Me
End Sub

Private Sub lstANLines_Click()
    Dim lineID As Long
    
    'Get SpoolHistory, Sensitivity data for this spool
    
'    lineID = IIf(IsNull(lstANLines.Column(0)), 0, lstANLines.Column(0))
    lineID = lstANLines.Value
    
    If lineID > 0 Then
        Call GetMSLData.GetSpools(, lineID)
        Call GetSpoolHistory(, lineID)
    End If
    
    Call ClearListBoxSelection(lstANLines)
    
    'Close form
    Unload Me
    
End Sub

Private Sub lstMLLines_Click()
    Dim lineID As Long
    
    'Get SpoolHistory, Sensitivity data for this spool
    
'    lineID = IIf(IsNull(lstMLLines.Column(0)), 0, lstMLLines.Column(0))
    lineID = lstMLLines.Value
    
    If lineID > 0 Then
        Call GetMSLData.GetSpools(, lineID)
        Call GetSpoolHistory(, lineID)
    End If
    
    Call ClearListBoxSelection(lstMLLines)
    
    'Close form
    Unload Me
    
End Sub

'Private Sub btnGetSpool_Click()
'    SpoolLocID As Long
'
'    If Not IsNull(txtSpool) Then
'        SpoolLocID = CurrentSpoolLocIDfromSpool(txtSpool)
'    End If
'
'    If SpoolLocID = 0 Then
'        MsgBox "Spool not found" & Chr(10) & Chr(10) & _
'               "Make sure the spool number is enetered correctly." & Chr(10) & _
'               "This spool may not be a part of the current Line configuration " & Chr(10) & _
'               "and therefore not in service."
'    End If
'
'End Sub
