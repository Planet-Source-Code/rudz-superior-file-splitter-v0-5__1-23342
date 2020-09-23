Attribute VB_Name = "modFileExists"
' by Rudy Alex Kohn (i think :)
' Returns TRUE if file exists, FALSE if not

Option Explicit

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Private Const OFS_MAXPATHNAME = 128
Private Const OF_EXIST = &H4000
Private Const FILE_NOT_FOUND = 2

Private Type OFSTRUCT
  cBytes As Byte
  fFixedDisk As Byte
  nErrCode As Integer
  Reserved1 As Integer
  Reserved2 As Integer
  szPathName(OFS_MAXPATHNAME) As Byte
End Type

Function FileExists(FileName As String) As Long
  Dim RetCode As Long
  Dim OpenFileStructure As OFSTRUCT

  RetCode = OpenFile(FileName$, OpenFileStructure, OF_EXIST)

  Select Case OpenFileStructure.nErrCode
    Case FILE_NOT_FOUND
        FileExists = False
    Case Else
        FileExists = True
  End Select

End Function

