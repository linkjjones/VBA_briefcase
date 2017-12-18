Attribute VB_Name = "DataDownloadExample"
Option Explicit

Public Sub WeldType()
    Dim Col                 As Long
    Dim ws                  As Worksheet
    Dim btmRow              As Long
    Dim ClearRange          As Range
    Dim sql                 As String
    Dim rs                  As New ADODB.Recordset
    Dim CheckCol            As Long
    
    Set ws = ActiveWorkbook.Worksheets("Control")
 
    ActiveSheet.Unprotect
    Application.StatusBar = "Downloading Weld Types..."
    
    Col = XLFunc.HeaderCol(ws, "WeldTypeID")
    CheckCol = Col - 1
    btmRow = XLFunc.GetLastRow(ws, Col + 1, DataStartRow)
   
    'Clear data
    Set ClearRange = ws.Range(ws.Cells(DataStartRow, Col - 1), ws.Cells(btmRow + 1, Col + 2))
    ClearRange.Clear
    ClearRange.ClearFormats
    
    'Connect to DB
    Call DBConnection.Connect
    
    'Copy from rs
    sql = "Select ID, TypeName, RTPercent " & _
          "From WeldType " & _
          "Order By SummaryOrder;"
    rs.Open sql, DBCON, 1, 3
    If rs.RecordCount > 0 Then
        ws.Cells(DataStartRow, Col).CopyFromRecordset rs
    End If
    rs.Close
    
    'Disconnect DB
    Call DBConnection.Disconnect
    
    btmRow = XLFunc.GetLastRow(ws, Col, DataStartRow)
    
    'Name the WeldTypeSelect range
    ws.Range(ws.Cells(DataStartRow, CheckCol), ws.Cells(btmRow + 1, CheckCol)).Name = "WeldTypeSelect"
    
    'Format
    'Update Column
    XLFunc.ValidateRange ws.Range("WeldTypeSelect"), "Update", False
    
    XLFunc.FormatRangeWithLines ws.Range(ws.Cells(DataStartRow, Col - 1), ws.Cells(btmRow + 1, Col + 2))
    
    'Whole data set
    btmRow = XLFunc.GetLastRow(ws, Col, DataStartRow)
    With ws.Range(ws.Cells(DataStartRow, Col - 1), ws.Cells(btmRow + 1, Col + 2))
        .HorizontalAlignment = xlLeft
    End With
    
    'ID Col
    With ws.Columns(Col)
        .Font.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
        .Font.Size = 8
    End With
    
    'Format Percent Col
    XLFunc.FormatPercent ws.Columns(Col + 2)
    
    Application.StatusBar = ""
    
End Sub

