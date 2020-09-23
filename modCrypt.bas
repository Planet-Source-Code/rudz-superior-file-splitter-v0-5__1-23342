Attribute VB_Name = "modCrypt"
' Otiginal by : Rami.Rasanen@p1.f73.n369.z1.fidonet.org
' Optimization by : Rudy Alex Kohn [vader@earthcorp.com]

Option Explicit

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private h As String, i As Integer, j As String, a As Integer

Sub EnCrypt(Pass As String, StringToEncrypt As String, KeyName As String, Path As String, SplitInfo As String)

    Dim Strg As String

    Strg = StringToEncrypt
    Call Crypt(Pass, Strg)

    h = vbNullString

    For i = 1 To Len(Strg)
        j = Hex$(Asc(Mid$(Strg, i, 1)))
        If Len(j) = 1 Then j = "0" + j
        h = h + j
    Next

    '   Store the LENGTH of the password string as 2 bytes and concatenate
    h = Format$(Len(h), "00") + h

    WritePrivateProfileString "File Splitter", KeyName, h, Path & SplitInfo

End Sub

Function DeCrypt(Pass As String, KeyName As String, Path As String, SplitInfo As String) As String
    Dim Strg As String

    ' To read it back in,
    h = Space$(80)
    GetPrivateProfileString "File Splitter", KeyName, "Unknown", h, Len(h), Path & SplitInfo

    h = Mid$(h, 3, Val(Left$(h, 2)))

    Strg = vbNullString
    For i = 1 To Len(h) Step 2
        j = Mid$(h, i, 2)
        Strg = Strg + Chr$(Val("&H" + j))
    Next

    Crypt Pass, Strg
    DeCrypt = Strg
End Function

Function Crypt(Password As String, Strg As String)
    Dim b As String
    a = 1
    For i = 1 To Len(Strg)
        b = Asc(Mid$(Password, a, 1))
        a = a + 1
        If a > Len(Password) Then a = 1
        Mid$(Strg, i, 1) = Chr$(Asc(Mid$(Strg, i, 1)) Xor b)
    Next
End Function
