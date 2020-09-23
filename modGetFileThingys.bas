Attribute VB_Name = "modFileThingy"
'made in a hurry by Rudy Alex Kohn
'rudyalexkohn@hotmail.com
'www.vaderdk.subnet.dk
Option Explicit

Function GetFileName(FileName As String) As String
'returns filename.ext from drive:\path\path\etc\filename.ext
    Dim i As Integer
    Dim tmp As String
    GetFileName = FileName
    For i = 1 To Len(FileName)
        tmp = Right$(FileName, i)
        If Left$(tmp, 1) = "\" Then
            GetFileName = Mid$(tmp, 2)
            Exit For
        End If
    Next
End Function

Function GetFileExtension(FileName As String, Optional LowerCase As Boolean = True) As String
' Returns .ext of filename.ext. If lowercase = true (default) then it will be _
  converted to small chars
    Dim i As Integer
    GetFileExtension = FileName     ' Just in case there is no "." in the file
    For i = 1 To Len(FileName)
        If Mid$(FileName, Len(FileName) - i, 1) = "." Then
            GetFileExtension = Mid$(FileName, Len(FileName) - i)
            Exit For
        End If
    Next
    If (LowerCase) Then GetFileExtension = LCase$(GetFileExtension)
End Function

Function GetFileNoExtension(FileName As String) As String
' Returns filename from filename.ext
    Dim i As Integer
    GetFileNoExtension = FileName     ' Just in case there is no "." in the file
    For i = 1 To Len(FileName)
        If Mid$(FileName, Len(FileName) - i, 1) = "." Then
            GetFileNoExtension = Mid$(FileName, 1, Len(FileName) - (i + 1))
            Exit For
        End If
    Next
End Function

Function GetFilePath(FileName As String, Optional IncludeDrive As Boolean = True) As String
' returns path. drive can be excluded if needed
    GetFilePath = FileName
    Dim i As Integer
    Dim str As String
    For i = 1 To Len(FileName)
        str = Right$(FileName, i)
        If Mid$(str, 1, 1) = "\" Then
            Dim iLenght As Integer
            If (IncludeDrive) Then iLenght = 1 Else iLenght = 4
            GetFilePath = Mid$(FileName, iLenght, Len(FileName) - i) & "\"
            Exit Function
        End If
    Next
End Function

Function GetDrive(FileName As String, Optional IncludeSlash As Boolean = False) As String
' returns lowercase drive ..
    Dim iLenght As Integer
    If (IncludeSlash) Then iLenght = 3 Else iLenght = 2
    GetDrive = LCase$(Left$(FileName, iLenght))
End Function
