Attribute VB_Name = "Email"
Option Explicit

Public Function SendMail(Email As String, Optional CC As String, _
                      Optional AttachmentStr_Arr As Variant) As Boolean
    Dim OutlApp As Object
    Dim IsCreated As Boolean
    Dim olEmail As Object
    Dim APPPath As String
    Dim attPath As Variant
    
    
    If Email = "" Then
        Exit Function
    End If
    
    ' Use already open Outlook if possible
    On Error Resume Next
    ' Set OutlApp = GetObject(, "Outlook.Application")
    'this one opens outlook and makes it visible
    Set OutlApp = OpenOutlook.OutlookApp()
'    Set olEmail = OutlApp.mailitem
    'So this isn't needed
    ' If Err Then
    ' Set OutlApp = FindOLexe.OpenOutLookApp
    ' IsCreated = True
    ' End If
    ' OutlApp.Visible = True
    On Error GoTo 0
    
    If OutlApp Is Nothing Then
        Exit Function
    End If
    
    'create email object
    Set olEmail = OutlApp.CreateItem(olMailItem)
    With olEmail
        .to = Email
        'Test****
'         .to = "jeff@dataspeaks.ca; duane@dataspeaks.ca"
'            CC = ""
        '********
        If Not CC = "" Then
            .CC = CC
        End If
        '.BCC = ""
        .Subject = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5)
        '.HTMLBody = CreateHTMLBody.RangetoHTML(Range("EmailFormat"))
        .Body = ""
        
        'Attachments
        If VarType(AttachmentStr_Arr) = vbArray Then
            For Each attPath In AttachmentStr_Arr
                .attachments.Add attPath
            Next attPath
        ElseIf VarType(AttachmentStr_Arr) = vbString Then
            .attachments.Add AttachmentStr_Arr
        End If
        
        ' Try to send
        On Error GoTo 0
        On Error Resume Next
        .Send
    End With
    
    If Err Then
        GoTo SendError
    End If
    
    SendMail = True
    
Cleanup:
    
    On Error GoTo 0
    ' ' Quit Outlook if it was created by this code
    ' If IsCreated Then OutlApp.Quit
    
    ' Release the memory of object variable
    Set OutlApp = Nothing
    
    Exit Function
SendError:
    SendMail = False
    Resume Cleanup

End Function
