VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SBRowForm 
   Caption         =   "Small Bore Connections"
   ClientHeight    =   2670
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   7950
   OleObjectBlob   =   "SBRowForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SBRowForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddRow_Click()
If PointName = "" Then Exit Sub
SBRows.AddItem PointName
End Sub

Private Sub AddRowsToWS_Click()

r = ActiveCell.Row
rh = Rows(r).RowHeight
Rows(r).RowHeight = rh / SBRows.ListCount

For X = 0 To SBRows.ListCount - 1
    If X = 0 Then
        Cells(r, 1) = SBRows.List(X)
    Else
        Rows(r + X).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Rows(r).Copy
        Rows(r + X).PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
            , SkipBlanks:=False, Transpose:=False
        Cells(r + X, 1) = SBRows.List(X)
    End If

Next X
Range("S" & r & ":S" & r + SBRows.ListCount - 1).Merge
Range("T" & r & ":T" & r + SBRows.ListCount - 1).Merge
Range("U" & r & ":U" & r + SBRows.ListCount - 1).Merge
Unload Me

End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub RemoveRow_Click()
If SBRows.ListIndex = -1 Then Exit Sub
SBRows.RemoveItem SBRows.ListIndex
End Sub
