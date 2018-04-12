VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControllerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objcon_person As clsModelPerson
Private m_objcon_form As clsViewForm

Private Sub Class_Initialize()
    Set con_form = New clsViewForm
    Set con_person = New clsModelPerson
End Sub

Private Sub Class_Terminate()
    Set m_objcon_person = Nothing

    Set m_objcon_form = Nothing
End Sub

Public Property Get con_person() As clsModelPerson

    Set con_person = m_objcon_person

End Property

Public Property Set con_person(objcon_person As clsModelPerson)

    Set m_objcon_person = objcon_person

End Property

Public Property Get con_form() As clsViewForm

    Set con_form = m_objcon_form

End Property

Public Property Set con_form(objcon_form As clsViewForm)

    Set m_objcon_form = objcon_form

End Property

Sub form_get(person_id As Integer)
    If con_person.GetPerson(person_id) Then
        con_form.clear_form
        con_form.update_form con_person
    Else
        con_form.clear_form
    End If
End Sub

Sub form_update()
    con_person.UpdatePerson con_form.extract_current_person
End Sub
