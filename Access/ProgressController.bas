Attribute VB_Name = "ProgressController"
Option Compare Database
Option Explicit

Private frm As String
Dim img As Control

Public Sub Show(Show As Boolean)
    
    frm = "ProgressMeter"
    If Show Then
        DoCmd.OpenForm frm, acNormal
        'loop through images and hide all but img000
        For Each img In Forms(frm)
            If Left(img.Name, 3) = "img" Then
                If Left(img.Name, 3) = "img" Then
                    If img.Name = "img0" Then
                        img.Visible = True
                    Else
                        img.Visible = False
                    End If
                End If
            End If
        Next img
        DoEvents
        Forms(frm).SetFocus
    Else
        DoCmd.Close acForm, frm
    End If
End Sub

Public Sub UpdatePercent(Percentage As Double)
    Percentage = Round(Percentage, 2) * 100
    'Set percent on meter
    Forms(frm).lblPercent.Caption = Percentage & "%"
    
    'Round to the nearest 5%
    If Round(Percentage / 5) * 5 = Percentage Then
        Percentage = Round(Percentage / 5) * 5
        For Each img In Forms(frm)
            If Left(img.Name, 3) = "img" Then
                If img.Name = "img" & CStr(Percentage) Then
                    DoEvents
                    img.Visible = True
                Else
                    img.Visible = False
                End If
            End If
        Next img
    End If
End Sub
