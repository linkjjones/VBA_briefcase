VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WBWS 
   Caption         =   "Data Transfer From Template To XTool"
   ClientHeight    =   4650
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   6705
   OleObjectBlob   =   "WBWS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WBWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
CancelTransfer = True
Unload Me
End Sub

Private Sub TransferButton_Click()
Set XTwb = Workbooks(WorksheetList.List(WorksheetList.ListIndex, 0))
Set XTsh = Workbooks(WorksheetList.List(WorksheetList.ListIndex, 0)).Sheets(WorksheetList.List(WorksheetList.ListIndex, 1))
Unload Me
End Sub

Private Sub UserForm_Initialize()


For X = 1 To Workbooks.Count

For Y = 1 To Workbooks(X).Sheets.Count

    WorksheetList.AddItem
    WorksheetList.List(Z, 0) = Workbooks(X).Name
    WorksheetList.List(Z, 1) = Workbooks(X).Sheets(Y).Name
    Z = Z + 1
Next Y

Next X

End Sub
