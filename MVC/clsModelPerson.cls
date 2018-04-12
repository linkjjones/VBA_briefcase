VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModelPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_isr_no As Integer
Private m_sfirst_name As String
Private m_slast_name As String
Private m_iage As Integer
Private m_ssex As String

Public Property Get sr_no() As Integer

    sr_no = m_isr_no

End Property

Public Property Let sr_no(ByVal isr_no As Integer)

    m_isr_no = isr_no

End Property

Public Property Get first_name() As String

    first_name = m_sfirst_name

End Property

Public Property Let first_name(ByVal sfirst_name As String)

    m_sfirst_name = sfirst_name

End Property

Public Property Get last_name() As String

    last_name = m_slast_name

End Property

Public Property Let last_name(ByVal slast_name As String)

    m_slast_name = slast_name

End Property

Public Property Get age() As Integer

    age = m_iage

End Property

Public Property Let age(ByVal iage As Integer)

    m_iage = iage

End Property

Public Property Get sex() As String

    sex = m_ssex

End Property

Public Property Let sex(ByVal ssex As String)

    m_ssex = ssex

End Property

Public Function GetPerson(person_id As Integer) As Boolean
    Dim per As clsModelPerson, iCount As Integer
    GetPerson = False
    For iCount = 1 To 4

        With Sheet1.Range("D1")

            If .Offset(iCount, 0).Value = person_id Then
                Me.first_name = .Offset(iCount, 1).Value
                Me.last_name = .Offset(iCount, 2).Value
                Me.age = .Offset(iCount, 3).Value
                Me.sex = .Offset(iCount, 4).Value
                Me.sr_no = .Offset(iCount, 0).Value
                GetPerson = True
            End If

        End With
    Next iCount

End Function

Public Function UpdatePerson(person As clsModelPerson) As Boolean
    
    Dim iCount As Integer
    UpdatePerson = False
    
    For iCount = 1 To 4
        With Sheet1.Range("D1")
            If .Offset(iCount, 0).Value = person.sr_no Then
                .Offset(iCount, 1).Value = person.first_name
                .Offset(iCount, 2).Value = person.last_name
                .Offset(iCount, 3).Value = person.age
                .Offset(iCount, 4).Value = person.sex
                UpdatePerson = True
            End If
        End With
    Next iCount

End Function
