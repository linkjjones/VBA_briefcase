Attribute VB_Name = "modMakeMultiStepDirectory"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modMakeMultiStepDirectory
' By Chip Pearson, chip@cpearson.com   www.cpearson.com
'
' This module contains the MakeMultiStepDirectory function that allows
' you to create an entire series of directories, creating each parent
' directory before each child directory. This works on both UNC
' and Local folder specifications. Examples:
'   MakeMultiStepDirectory "C:\One\Two\Three"
'       This will create    C:\One      then
'                           C:\One\Two  then
'                           C:\One\Two\Three
'
'   MakeMultiStepDirectory "\\BlackCow\MainShare\One\Two\Three
'       This creates    \\BlackCow\MainShare\One        then
'                       \\BlackCow\MainShare\One\Two    then
'                       \\BlackCow\MainShare\One\Two\Three
'
' If any directory already exists, it is skipped until a non-extant
' directory name is encountered.
'
' Return values:
'        ErrSuccess
'           Operation was successful.
'        ErrRelativePath
'           The input PathSpec was not an absolute path.
'        ErrInvalidPathSpecification
'           Invalid path specification.
'        ErrDirectoryCreateError
'           Error creating directory.
'        ErrSpecIsFileName
'           FileSpec is a file name, not a folder name.
'        ErrInvalidCharactersInPath
'           Invalid characters found in PathSpec.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Required Reference:
'   Name: Scripting
'   Description: Microsoft Scripting Runtime
'   Typical Location: C:\Windows\SysWOW64\scrrun.dll
'   GUID:   {420B2830-E718-11CF-893D-00A0C9054228}
'   Major: 1    Minor: 0

Private Declare Function PathIsRelative Lib "Shlwapi" _
    Alias "PathIsRelativeA" (ByVal Path As String) As Long

Public Enum EMakeDirStatus
    ErrSuccess = 0
    ErrRelativePath
    ErrInvalidPathSpecification
    ErrDirectoryCreateError
    ErrSpecIsFileName
    ErrInvalidCharactersInPath
End Enum
Const MAX_PATH = 260

Function MakeMultiStepDirectory(ByVal PathSpec As String) As EMakeDirStatus
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MakeMultiStepDirectory
' This function creates a series of nested directories. The parent of
' every directory is create before a subdirectory is created, allowing a
' folder path specification of any number of directories (as long as the
' total length is less than MAX_PATH.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FSO As Scripting.FileSystemObject
Dim DD As Scripting.Drive
Dim B As Boolean
Dim Root As String
Dim DirSpec As String
Dim N As Long
Dim M As Long
Dim S As String
Dim Directories() As String
    
Set FSO = New Scripting.FileSystemObject
    
' ensure there are no invalid characters in spec.
On Error Resume Next
Err.Clear
S = Dir(PathSpec, vbNormal)
If Err.Number <> 0 Then
    MakeMultiStepDirectory = ErrInvalidCharactersInPath
    Exit Function
End If
On Error GoTo 0

' ensure we have an absolute path
B = CBool(PathIsRelative(PathSpec))
If B = True Then
    MakeMultiStepDirectory = ErrRelativePath
    Exit Function
End If

' if the directory already exists, get out with success.
If FSO.FolderExists(PathSpec) = True Then
    MakeMultiStepDirectory = ErrSuccess
    Exit Function
End If

' get rid of trailing slash
If Right(PathSpec, 1) = "\" Then
    PathSpec = Left(PathSpec, Len(PathSpec) - 1)
End If

' ensure we don't have a filename
N = InStrRev(PathSpec, "\")
M = InStrRev(PathSpec, ".")
If (N > 0) And (M > 0) Then
    If M > N Then
        ' period found after last slash
        MakeMultiStepDirectory = ErrSpecIsFileName
        Exit Function
    End If
End If

If Left(PathSpec, 2) = "\\" Then
    ' UNC -> \\Server\Share\Folder...
    N = InStr(3, PathSpec, "\")
    N = InStr(N + 1, PathSpec, "\")
    Root = Left(PathSpec, N - 1)
    DirSpec = Mid(PathSpec, N + 1)
Else
    ' Local or mapped -> C:\Folder....
    N = InStr(1, PathSpec, ":", vbBinaryCompare)
    If N = 0 Then
        MakeMultiStepDirectory = ErrInvalidPathSpecification
        Exit Function
    End If
    Root = Left(PathSpec, N)
    DirSpec = Mid(PathSpec, N + 2)
End If
Set DD = FSO.GetDrive(Root)
Directories = Split(DirSpec, "\")
DirSpec = DD.Path
For N = LBound(Directories) To UBound(Directories)
    DirSpec = DirSpec & "\" & Directories(N)
    If FSO.FolderExists(DirSpec) = False Then
        On Error Resume Next
        Err.Clear
        FSO.CreateFolder (DirSpec)
        If Err.Number <> 0 Then
            MakeMultiStepDirectory = ErrDirectoryCreateError
            Exit Function
        End If
    End If
Next N
MakeMultiStepDirectory = ErrSuccess

End Function

