Attribute VB_Name = "modIntiatior"
Option Explicit

Public controller As clsControllerForm

Sub form_getdata()
    controller.form_get Sheet1.Range("B1").Value
End Sub

Sub form_updatedata()
    controller.form_update
End Sub


