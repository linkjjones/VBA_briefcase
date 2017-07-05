Attribute VB_Name = "Security"
Option Compare Database
Option Explicit

Public Function IsPassword(strValue As String)
    If StrComp(strValue, DLookup("PW", "Pass"), vbBinaryCompare) = 0 Then
        IsPassword = True
    End If
End Function

Public Function ValidPassword(ByVal strPass As String) As Boolean
    
    If HasNumbersAndLettersOnly(strPass) Then
        If Len(strPass) > 7 Then
            ValidPassword = True
            Exit Function
        End If
    End If
    
    MsgBox "Password must contain letters and at least one number only."
    
End Function

Public Function HasNumbersAndLettersOnly(ByVal strValue As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim NonAlphaNumericCount As Integer
    
    strValue = UCase(strValue)
    For i = 1 To Len(strValue)
        'returns...
        '$ : String
        '% : Integer (Int32)
        '& : Long (Int64)
        '! : Single
        '# : Double
        '@ : Decimal
        ch = Mid$(strValue, i, 1)
        If ch < "A" Or ch > "Z" Then
            'this is not a letter
            If ch < "0" Or ch > "9" Then
                'this is not a number
                NonAlphaNumericCount = NonAlphaNumericCount + 1
            End If
        End If
    Next i
    
    If NonAlphaNumericCount = 0 Then
        HasNumbersAndLettersOnly = True
    End If
    
End Function

Public Function UpdatePassword(newPW As String)
    CurrentDb.Execute "Update Pass Set PW='" & newPW & "';"
End Function
