VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MinTPoint 
   Caption         =   "Minimum Thickness Band and Point"
   ClientHeight    =   2445
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   7845
   OleObjectBlob   =   "MinTPoint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MinTPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub FirstBand_Change()
ConstructedTML = TML & FirstBand & SecondBand & Point
End Sub

Private Sub Ok_Click()
ActiveCell = ConstructedTML.Caption
Unload Me
End Sub

Private Sub Point_Change()
ConstructedTML = TML & FirstBand & SecondBand & Point
End Sub

Private Sub SecondBand_Change()
ConstructedTML = TML & FirstBand & SecondBand & Point
End Sub

Private Sub UserForm_Initialize()
FirstBand.AddItem "A"
FirstBand.AddItem "B"
FirstBand.AddItem "C"
FirstBand.AddItem "D"
FirstBand.AddItem "E"
FirstBand.AddItem "F"
FirstBand.AddItem "G"

SecondBand.AddItem "A"
SecondBand.AddItem "B"
SecondBand.AddItem "C"
SecondBand.AddItem "D"
SecondBand.AddItem "E"
SecondBand.AddItem "F"
SecondBand.AddItem "G"

Point.AddItem "A"
Point.AddItem "B"
Point.AddItem "C"
Point.AddItem "D"
Point.AddItem "E"
Point.AddItem "F"
Point.AddItem "G"
Point.AddItem "H"

ConstructedTML.Caption = TML

End Sub
