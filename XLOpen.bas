Attribute VB_Name = "XLOpen"
Option Explicit

Public Function OpenWorkbook(FilePath As String, Visible As Boolean, _
                             Optional Password As String, _
                             Optional WriteMode As Boolean, _
                             Optional ShowOpenAlerts As Boolean) _
                             As Excel.Workbook
    On Error GoTo errHandler
    Dim XLApp As Excel.Application
    Dim XLBook As Excel.Workbook
    
    If Not fn.fileExists(FilePath, file) Then
        
    End If
    
    'Open Spreadsheet
    Set XLApp = CreateObject("Excel.application")

OpenXLBook:
    
    If Not ShowOpenAlerts Then
        XLApp.Application.DisplayAlerts = False
        XLApp.Application.AskToUpdateLinks = False
        XLApp.Application.DisplayAlerts = True
        XLApp.Application.AskToUpdateLinks = True
    End If
    
    If WriteMode Then
        If Password = "" Or IsNull(Password) Then
            Set XLBook = XLApp.Workbooks.Open(FilePath, , False, , , , True, , , True)
        Else
            Set XLBook = XLApp.Workbooks.Open(FilePath, , False, , Password, , True, , , True)
        End If
    Else
        If Password = "" Or IsNull(Password) Then
            Set XLBook = XLApp.Workbooks.Open(FilePath, , True, , , , True)
        Else
            Set XLBook = XLApp.Workbooks.Open(FilePath, , True, , Password, , True)
        End If
    End If
    
    XLApp.Application.ActiveWorkbook.UpdateLinks = xlUpdateLinksAlways
    
    XLApp.Visible = Visible
    
    Set OpenWorkbook = XLBook
    
GCleanUp:
    Set XLApp = Nothing
    Set XLBook = Nothing
    
    
Exit Function
errHandler:
If Err.Number = 1004 Then
    MsgBox "Cannot access file: " & Chr(10) & FilePath, vbInformation
    XLApp.Quit
End If



End Function

Public Sub CloseWorkbook(XLBook As Excel.Workbook, Optional SaveWB As Boolean)
    Dim XLApp As Excel.Application
    
    Set XLApp = XLBook.Application
    If Not XLBook Is Nothing Then
        If SaveWB Then
            XLBook.Save
        Else
            XLBook.Saved = True
        End If
    End If
    XLApp.Quit
    Set XLApp = Nothing
    
End Sub

