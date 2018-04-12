VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BandViewerForm 
   Caption         =   "Band Viewer"
   ClientHeight    =   6276
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   19560
   OleObjectBlob   =   "BandViewerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BandViewerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub UserForm_Activate()
With Me
        'This will create a vertical scrollbar
        .ScrollBars = fmScrollBarsVertical
        
        'Change the values of 2 as Per your requirements
        .ScrollHeight = .InsideHeight * 2.2
        '.ScrollWidth = .InsideWidth * 9
    End With

End Sub

