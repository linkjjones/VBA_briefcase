Attribute VB_Name = "ConnectDatabase"
Option Explicit
Public DBCON As Object
Public WHCON As Object

'OpendatabaseConnection
Public Function ConnectDB() As Object
    Dim strDBPath   As String
    Dim colFilePath As Long
    Dim rowfilepath As Long
    Dim strCon      As String
    Dim wsCtl       As Worksheet
    
    Set wsCtl = Worksheets("Control")
    
    colFilePath = HeaderCell("File Paths", 3, wsCtl, LastCol(wsCtl, 3)).Column
    rowfilepath = FindRowInColumn(wsCtl, colFilePath, 4, LastRow(wsCtl, colFilePath), "BackEnd")
    strDBPath = wsCtl.Cells(rowfilepath, colFilePath + 1).Value
    
    Set DBCON = CreateObject("ADODB.connection")
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=" & strDBPath & ";" & _
             "Jet OLEDB:Engine Type=5;" & _
             "Persist Security Info=False;"
    On Error GoTo 2003
    DBCON.Open strCon

    Exit Function
2003:
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & strDBPath & ";" & _
             "Persist Security Info=False;"
    On Error GoTo Error
    DBCON.Open strCon

    Exit Function
Error:
    Call DisconnectDB
End Function

Public Function DisconnectDB()
    On Error Resume Next
    DBCON.Close
    Set DBCON = Nothing
    On Error GoTo 0
End Function

Public Function TableAvailable(TableName As String, FieldName As String, _
                                  Optional CloseOnCancel = False) As Boolean
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim DesktopPath As String
    Dim msg As String
    Dim TimeLimit
    Dim startTime
    TimeLimit = 3
    msg = "Database is currently unavailable." & Chr(10) & _
           "This could be due to a process running" & Chr(10) & _
           "or your network connection has been lost." & Chr(10) & Chr(10) & _
           "Try again?"
Try:
    startTime = Timer
    Application.StatusBar = True
    Application.StatusBar = "Checking database availability..."
    Do While Timer < startTime + TimeLimit
        Do While Not TableAvailable
'            'reset error
'            Error = 0
            On Error Resume Next
            sql = "Select " & FieldName & " From [" & TableName & "];"
            rs.Open sql, DBCON, 1, 3
            If Err = 0 Then
                rs.Close
                Set rs = Nothing
                TableAvailable = True
                Application.StatusBar = False
                Exit Function
            Else    'check if it has timed out
                If Timer >= startTime + TimeLimit Then
                    Exit Do
                Else
                    MsgBox "This Workbook will be saved to your desktop and closed." & Chr(10) & _
                           "Please open it at a later time and run the upload again.", vbInformation, _
                           "Try again later"
                    'Save and close
                    DesktopPath = "C:\Users\" & Environ("USERNAME") & "\Desktop"
                    DesktopPath = DesktopPath & "Open and Upload - " & Format(Date, "YYYY-MMM-D HH-MM-SS") & ".xlsm"
                    Call SaveToDesktop(DesktopPath, True)
                    ActiveWorkbook.Close False
                End If
            End If
        Loop
    Loop
    
    If MsgBox(msg, vbRetryCancel + vbCritical, "Database connection failed") = vbRetry Then
        'retry
        GoTo Try
    Else
        'cancel
        If CloseOnCancel Then
            Application.Quit
        Else
            'some other option
        End If
    End If
    
End Function

Private Sub SaveToDesktop(FullFileName As String, _
                          Optional RecommendReadOnly As Boolean = False)
    
    Application.DisplayAlerts = False
    Call DisconnectDB
    
    ActiveWorkbook.SaveAs FullFileName, , , , RecommendReadOnly
    
    Application.DisplayAlerts = True
    
End Sub

