VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This sub-procedure is to update the form data

Sub update_form(person As clsModelPerson)
    With Sheet1.Range("B1")
        .Offset(1, 0).Value = person.first_name
        .Offset(2, 0).Value = person.last_name
        .Offset(3, 0).Value = person.age
        .Offset(4, 0).Value = person.sex
    End With
End Sub

'This procedure is for clearing the form data
Sub clear_form()
    With Sheet1.Range("B1")
        .Offset(1, 0).ClearContents
        .Offset(2, 0).ClearContents
        .Offset(3, 0).ClearContents
        .Offset(4, 0).ClearContents
    End With
End Sub

Function extract_current_person() As clsModelPerson

    Set extract_current_person = New clsModelPerson
    With extract_current_person
        .first_name = Sheet1.Range("B1").Offset(1, 0).Value
        .last_name = Sheet1.Range("B1").Offset(2, 0).Value
        .age = Sheet1.Range("B1").Offset(3, 0).Value
        .sex = Sheet1.Range("B1").Offset(4, 0).Value
        .sr_no = Sheet1.Range("B1").Offset(0, 0).Value
    End With

End Function
