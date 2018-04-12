Attribute VB_Name = "Module5"
' Access the GetUserNameA function in advapi32.dll and
' call the function GetUserName.
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
 (ByVal lpBuffer As String, nSize As Long) As Long
 
' Main routine to retrieve user name.
Function GetLogonName() As String
 
 ' Dimension variables
 Dim lpBuff As String * 255
 Dim ret As Long
 
 ' Get the user name minus any trailing spaces found in the name.
 ret = GetUserName(lpBuff, 255)
 
 If ret > 0 Then
 GetLogonName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
 Else
 GetLogonName = vbNullString
 End If
 
End Function


