VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PictureViewer 
   Caption         =   "Picture Viewer"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   13245
   OleObjectBlob   =   "PictureViewer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PictureViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ScrollBar1_Change()
PictureViewer.Zoom = ScrollBar1.Value
    PictureViewer.ZoomLevel = ScrollBar1.Value

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
ScrollBar1.Max = 100
    ScrollBar1.Min = 50
    ScrollBar1.Value = 100
    PictureViewer.Zoom = 100
End Sub
