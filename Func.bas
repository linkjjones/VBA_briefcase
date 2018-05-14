Attribute VB_Name = "Func"
Option Explicit

Public Enum PathType
    Directory = 0
    file = 1
End Enum

Public Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayEmpty
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is really the reverse of IsArrayAllocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim LB As Long
    Dim UB As Long

    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I
        ' cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occasions, LBound is 0 and
        ' UBound is -1.
        ' To accommodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(Arr)
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function

Public Sub SayThis(Sentence As String)
    Dim s As Object
    Dim vol As Long
    
    Set s = CreateObject("SAPI.SpVoice")
    
    'Get current volume
    vol = s.volume
    'set higher volume
    s.volume = 100
    'Say it
    s.Speak Sentence
    'Set volume to original level
    s.volume = vol
    'Cleanup
    Set s = Nothing
    
End Sub
    
Public Function Friendly(handleDblQuotes As String) As String
    Friendly = Replace(handleDblQuotes, Chr(34), Chr(34) & Chr(34))
End Function
    
Public Function GetFile(Optional DialogTitle As String, _
                        Optional FileDescription As String, _
                        Optional FileExtension As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        If Not DialogTitle = "" Then
            .title = DialogTitle
        End If
        .Filters.Clear
        If Not FileExtension = "" Then
            .Filters.Add FileDescription, FileExtension, 1
        End If
        .InitialFileName = Application.ActiveWorkbook.Path
        
        If fd.Show = -1 Then
            GetFile = .SelectedItems(1)
        End If
    End With

End Function

Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function RemoveAmpersand(str As String) As String
    str = Replace(str, " & ", " ")
    str = Replace(str, "&", " ")
    RemoveAmpersand = str
End Function

Public Function TransposeArray(myarray As Variant) As Variant
'from https://bettersolutions.com/vba/arrays/transposing.htm
    Dim X As Long
    Dim Y As Long
    Dim Xupper As Long
    Dim Yupper As Long
    Dim tempArray As Variant
    Xupper = UBound(myarray, 2)
    Yupper = UBound(myarray, 1)
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = myarray(Y, X)
        Next Y
    Next X
    TransposeArray = tempArray
End Function

Public Function fileExists(file As String, fType As PathType) As Boolean
    If Not file = "" Then
        If fType = Directory Then
            fileExists = Dir(file, vbDirectory) <> ""
        Else
            fileExists = Dir(file) <> ""
        End If
    End If
End Function

Public Function GetPath(DialogTitle As String, namedRange As String, _
                      pType As PathType, _
                      Optional fileTypeDesc As String, _
                      Optional fileType As String)
    Dim fd As FileDialog
    Dim sPath As String
    Dim ws As Worksheet
    Set ws = Sheets("Control")
    If pType = file Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    
    With fd
        .title = DialogTitle
        If Not fileType = "" Then
            .Filters.Clear
            .Filters.Add fileTypeDesc, fileType, 1
        End If
        .InitialFileName = Application.ActiveWorkbook.Path
        
        If fd.Show = -1 Then
            sPath = .SelectedItems(1)
        End If
    End With
    
    ws.Range(namedRange) = sPath
End Function

Public Function mergeSort(c As Collection, Optional uniq = True) As Collection

    Dim i As Long, xMax As Long, tmp1 As Collection, tmp2 As Collection, xOdd As Boolean

    Set tmp1 = New Collection
    Set tmp2 = New Collection

    If c.Count = 1 Then
        Set mergeSort = c
    Else

        xMax = c.Count
        xOdd = (c.Count Mod 2 = 0)
        xMax = (xMax / 2) + 0.1     ' 3 \ 2 = 1; 3 / 2 = 2; 0.1 to round up 2.5 to 3

        For i = 1 To xMax
            tmp1.Add c.Item(i) & "" 'force numbers to string
            If (i < xMax) Or (i = xMax And xOdd) Then tmp2.Add c.Item(i + xMax) & ""
        Next i

        Set tmp1 = mergeSort(tmp1, uniq)
        Set tmp2 = mergeSort(tmp2, uniq)

        Set mergeSort = merge(tmp1, tmp2, uniq)

    End If
    
End Function

Private Function merge(c1 As Collection, c2 As Collection, _
                       Optional ByVal uniq As Boolean = True) As Collection

    Dim tmp As Collection
    Set tmp = New Collection

    If uniq = True Then On Error Resume Next    'hide duplicate errors

    Do While c1.Count <> 0 And c2.Count <> 0
        If c1.Item(1) > c2.Item(1) Then
            If uniq Then tmp.Add c2.Item(1), c2.Item(1) Else tmp.Add c2.Item(1)
            c2.Remove 1
        Else
            If uniq Then tmp.Add c1.Item(1), c1.Item(1) Else tmp.Add c1.Item(1)
            c1.Remove 1
        End If
    Loop

    Do While c1.Count <> 0
        If uniq Then tmp.Add c1.Item(1), c1.Item(1) Else tmp.Add c1.Item(1)
        c1.Remove 1
    Loop
    Do While c2.Count <> 0
        If uniq Then tmp.Add c2.Item(1), c2.Item(1) Else tmp.Add c2.Item(1)
        c2.Remove 1
    Loop
    On Error GoTo 0

    Set merge = tmp

End Function
