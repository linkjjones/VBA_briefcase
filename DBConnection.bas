Attribute VB_Name = "DBConnection"
Option Explicit
Public DBCON As Object
Public WHCON As Object

'OpendatabaseConnection
Public Function Connect(Optional QuitOnError As Boolean) As Boolean
    Dim strCon      As String
    Dim wsCtl       As Worksheet
    Dim strDBPath   As String
    Dim LiveDevCol  As Long
    Dim BackendType As String
    Dim PathRow     As Long
    Dim Msg         As String
    
    On Error GoTo ConErr
    
    Set wsCtl = Sheets("Control")
    BackendType = IIf(CBool(wsCtl.Range("ControlScaffold").Value), "Dev", "Live")
    
    LiveDevCol = XLFunc.HeaderCol(wsCtl, "Live / Dev")
    PathRow = XLFunc.FindRowInColumn(wsCtl, LiveDevCol, DataStartRow, XLFunc.GetLastRow(wsCtl, LiveDevCol, DataStartRow), BackendType)
    strDBPath = wsCtl.Cells(PathRow, LiveDevCol + 1)
    
    Set DBCON = CreateObject("ADODB.connection")
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
             "Data Source=" & strDBPath & ";" & _
             "Jet OLEDB:Engine Type=5;" & _
             "Persist Security Info=False;"
    DBCON.Open strCon
    'check if usser has Read/Write access
    DBCON.Execute "Update ConnTest Set TestField=True;"
    Connect = True
    
Exit Function
ConErr:
    Msg = "Excel cannot access the database." & Chr(10) & _
          "You may need to request LAN access to the QC folders."
    If QuitOnError Then
        Msg = Msg & Chr(10) & Chr(10) & _
              "Excel will now close this Welder Percent Log workbook."
    End If
    MsgBox Msg, vbInformation, "You do not have access to QC folders."
    DBConnection.Disconnect
    If QuitOnError Then
        ActiveWorkbook.Close False
    End If
        
End Function

Public Function Disconnect()
    On Error Resume Next
    DBCON.Close
    Set DBCON = Nothing
End Function

Public Function TableAvailable(TableName As String, FieldName As String, _
                                  Optional CloseOnCancel = False) As Boolean
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim DesktopPath As String
    Dim Msg As String
    Dim TimeLimit
    Dim startTime
    TimeLimit = 3
    Msg = "Database is currently unavailable." & Chr(10) & _
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
    
    If MsgBox(Msg, vbRetryCancel + vbCritical, "Database connection failed") = vbRetry Then
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
    Call Disconnect
    
    ActiveWorkbook.SaveAs FullFileName, , , , RecommendReadOnly
    
    Application.DisplayAlerts = True
    
End Sub

