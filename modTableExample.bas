Attribute VB_Name = "modWeldProc"
Option Explicit

Public Function GetID(ProcNumber As String) As Long
    'Get the ID from the passed ProcNumber
    'if it doesn't exist, return 0
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    'validation
    If ProcNumber = "" Then
        Exit Function
    End If
    
    Call DBConnection.Connect
    
    sql = "Select ID " & _
          "From WeldProc " & _
          "Where ProcNumber=""" & ProcNumber & """;"
    rs.Open sql, DBCON, 1, 3
    If rs.RecordCount > 0 Then
        GetID = rs!ID
    End If
    rs.Close
    
    Call DBConnection.Disconnect
    
    Set rs = Nothing
    
End Function

Public Function add(ProcNumber As String) As Long
    'check if ProcNumber exists
    'Add if it doesn't and return ID
    Dim rs As New ADODB.Recordset
    
    'validation
    If ProcNumber = "" Then
        Exit Function
    End If
    
    If modWeldProc.GetID(ProcNumber) > 0 Then
        Exit Function
    End If
    
    Call DBConnection.Connect
    
    rs.Open "WeldProc", DBCON, 1, 3
    rs.AddNew
    rs!ProcNumber = ProcNumber
    rs.update
    add = rs!ID
    rs.Close
    Set rs = Nothing
    
    Call DBConnection.Disconnect
    
End Function

Public Function update(ID As Long, ProcNumber As String) As Boolean
    Dim sql As String
    
    'Everything should be updateable
'    'Make sure the ProcNumber does not exists
'    If modWeldProc.GetID(ProcNumber) > 0 Then
'        Exit Function
'    End If
    
    If ID < 1 Then
        Exit Function
    End If
    
    Call DBConnection.Connect
    
    'update ProcNumber
    sql = "Update WeldProc " & _
          "Set ProcNumber=""" & ProcNumber & """ " & _
          "Where ID=" & ID & ";"
    DBCON.Execute sql
    update = True
    
    Call DBConnection.Disconnect
    
End Function

Public Sub Delete(ID As Long)
    'Chek if ID is used in Log
    'if not delete it
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    Call DBConnection.Connect
    
    sql = "Select ProcID " & _
          "From Log " & _
          "Where ProcID=" & ID & ";"
    rs.Open sql, DBCON, 1, 3
    If rs.RecordCount > 0 Then
        rs.Close
        GoTo Cleanup
    End If
    rs.Close
    
    sql = "Delete * From WeldProc " & _
          "Where ID=" & ID & ";"
    DBCON.Execute sql
    
Cleanup:
Call DBConnection.Disconnect
Set rs = Nothing

End Sub

