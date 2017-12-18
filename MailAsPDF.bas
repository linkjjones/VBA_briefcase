Attribute VB_Name = "MailAsPDF"
Option Explicit

Public Sub MailPDF()
  Dim IsCreated As Boolean
  Dim i As Long
  Dim PdfFile As String, Title As String
  Dim OutlApp As Object
  Dim Addresses As String
  Dim WS As Worksheet, DataRow As Integer, ctlCol As Integer, CCAddress As String
  
  If Left(Application.UserName, 5) <> "Jones" Then
    MsgBox "You are not authorized to send this as an email from this button."
    Exit Sub
  End If
  
  Set WS = Worksheets("Control")
  
  ' Not sure for what the Title is
  Title = "Sensitivities"
 
  ' Define PDF filename
  PdfFile = ActiveWorkbook.FullName
  i = InStrRev(PdfFile, ".")
  If i > 1 Then PdfFile = Left(PdfFile, i - 1)
  PdfFile = PdfFile & "_" & ActiveSheet.Name & ".pdf"
  
  ' Export activesheet as PDF
  With ActiveSheet
    .ExportAsFixedFormat Type:=xlTypePDF, Filename:=PdfFile, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
  End With
 
  ' Use already open Outlook if possible
  On Error Resume Next
  Set OutlApp = GetObject(, "Outlook.Application")
  If Err Then
    Set OutlApp = CreateObject("Outlook.Application")
    IsCreated = True
  End If
  OutlApp.Visible = True
  On Error GoTo 0
 
  ' Prepare e-mail with PDF attachment
  With OutlApp.CreateItem(0)
   
   Addresses = ""
   CCAddress = ""
   
   ctlCol = 10: DataRow = 6
    'Get the recipients (to:)
    Do While WS.Cells(DataRow, ctlCol) <> ""
        Addresses = Addresses & WS.Cells(DataRow, ctlCol) & "; "
        DataRow = DataRow + 1
    Loop
    
    'Take off the last ", "
    Addresses = Left(Addresses, Len(Addresses) - 2)
    
    ctlCol = 11: DataRow = 6
    'Get the recipients (to:)
    Do While WS.Cells(DataRow, ctlCol) <> ""
        CCAddress = CCAddress & WS.Cells(DataRow, ctlCol) & "; "
        DataRow = DataRow + 1
    Loop
    
    'Take off the last ", "
    CCAddress = Left(CCAddress, Len(CCAddress) - 2)
    
    ' Prepare e-mail
    .Subject = Title
    .To = Addresses ' <-- Put email of the recipient here
'    .To = "jones.jeffrey@syncrude.com ' <-- Put email of the recipient here"
    .CC = CCAddress '<-- Put email of 'copy to' recipient here"
    .Body = "Hi," & vbLf & vbLf _
          & "The Sensitivity List is attached in PDF format." & vbLf & vbLf _
          & "Regards," & vbLf _
          & Application.UserName & vbLf & vbLf
    .Attachments.Add PdfFile
   
    ' Try to send
    On Error Resume Next
    .Send
    Application.Visible = True
    If Err Then
      MsgBox "E-mail was not sent", vbExclamation
    Else
      MsgBox "E-mail successfully sent", vbInformation
    End If
    On Error GoTo 0
   
  End With
 
  ' Delete PDF file
  Kill PdfFile
 
  ' Quit Outlook if it was created by this code
  If IsCreated Then OutlApp.Quit
 
  ' Release the memory of object variable
  Set OutlApp = Nothing
 
    
End Sub

