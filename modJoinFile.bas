Attribute VB_Name = "modJoin"
Option Explicit

Private Huffman As clsHuffman

'Declare the variable Bytes as of ChunkSize type
Dim Bytes As ChunkSize

Function JoinFile(FileName As String, Extension As String, Compress As Integer, NumOfSegments As Integer, DestinationPath As String) As Integer
' Filename      = The original file
' Extension     = Original Extension
' Compress      = Compressed state
' NumOfSegments = Number of files to join



  On Error GoTo ErrorHandler

    
  'Make sure the settings are correct
  Dim ErrorCode As Integer
  If NumOfSegments > 999 Then                               ' Ensure that the segment size is valid
    ErrorCode = 2
    GoTo ErrorHandler
  End If

  Dim SourceBytes         As Long
  Dim SourceFile          As String
  Dim DestinationFile     As String
  Dim SegmentNumber       As Integer
  Dim BytesDone           As Long
  Dim FPath               As String
  Dim FName               As String
  Dim FNameNoExt          As String
  Dim RemainingBytes      As Long
  Dim FileSize            As Long



  FName = GetFileName(FileName)                             ' Extract the file name
  FPath = DestinationPath                                   ' Extract the path name
  FNameNoExt = GetFileNoExtension(FName)                    '   File name without extension

    
  '   Open the source file for binary write depending if file is compressed
  Select Case Compress
  Case 1
    Open FPath & FNameNoExt & ".tmp" For Binary Access Write As #2 Len = 1
  Case Else
    Open FPath & FNameNoExt & Extension For Binary Access Write As #2 Len = 1
  End Select

  For SegmentNumber = 1 To NumOfSegments

  ' Compose the file name of the next file to be read (file segment)
    Select Case SegmentNumber
      Case Is < 10
        SourceFile = FPath & FNameNoExt & ".00" & CStr(SegmentNumber)
      Case 10 To 99
        SourceFile = FPath & FNameNoExt & ".0" & CStr(SegmentNumber)
      Case 100 To 999
        SourceFile = FPath & FNameNoExt & "." & CStr(SegmentNumber)
    End Select

    SourceBytes = FileLen(SourceFile)         ' Get file lenght
    FileSize = SourceBytes

    '   Create the new file segment and open it for binary read
    Open SourceFile For Binary Access Read As #1 Len = 1

      BytesDone = 0
        
      ' Check whether the remaining bytes to process in the source file are less than Segment bytes
      Select Case SourceBytes - BytesDone
      Case Is < FileSize
        RemainingBytes = SourceBytes - BytesDone
      Case Else
        RemainingBytes = FileSize
      End Select

      frmProgress.lblStatus = "Writing..."
      frmProgress.lblCurrentFile = GetFileName(SourceFile)

      ' Read bytes from the source file and write them to the destination file (the current segment file) _
        Depending on the remaining bytes to read and write, the routine below will read the largest possible _
        chunk of data
      Do
        Select Case RemainingBytes
        Case Is >= 12000
          Get #1, , Bytes.S12000                    ' Read 12000 bytes of data from the source file
          Put #2, , Bytes.S12000                    ' Write 12000 bytes of data to the destination file
          RemainingBytes = RemainingBytes - 12000   ' Decrease the number of remaining bytes by 12000
          BytesDone = BytesDone + 12000             ' Update the bytes done counter
        Case 6000 To 11999
          Get #1, , Bytes.S6000
          Put #2, , Bytes.S6000
          RemainingBytes = RemainingBytes - 6000
          BytesDone = BytesDone + 6000
        Case 3000 To 5999
          Get #1, , Bytes.S3000
          Put #2, , Bytes.S3000
          RemainingBytes = RemainingBytes - 3000
          BytesDone = BytesDone + 3000
        Case 1500 To 2999
          Get #1, , Bytes.S1500
          Put #2, , Bytes.S1500
          RemainingBytes = RemainingBytes - 1500
          BytesDone = BytesDone + 1500
        Case 500 To 1499
          Get #1, , Bytes.S500
          Put #2, , Bytes.S500
          RemainingBytes = RemainingBytes - 500
          BytesDone = BytesDone + 500
        Case 100 To 499
          Get #1, , Bytes.S100
          Put #2, , Bytes.S100
          RemainingBytes = RemainingBytes - 100
          BytesDone = BytesDone + 100
        Case 25 To 99
          Get #1, , Bytes.S25
          Put #2, , Bytes.S25
          RemainingBytes = RemainingBytes - 25
          BytesDone = BytesDone + 25
        Case 5 To 24
          Get #1, , Bytes.S5
          Put #2, , Bytes.S5
          RemainingBytes = RemainingBytes - 5
          BytesDone = BytesDone + 5
        Case 1 To 4
          Get #1, , Bytes.S1
          Put #2, , Bytes.S1
          RemainingBytes = RemainingBytes - 1
          BytesDone = BytesDone + 1
        Case 0
          ' When the loop enters here, the segment bytes are completed. _
            Close the segment file and exit the loop
          Close 1
          DoEvents
          Exit Do
        End Select
        ' Update the percent control on the form
        DrawPercent frmProgress.picProgress, Int((BytesDone / SourceBytes) * 100), , vbWhite
        ' Yield to windows and other processes to do their jobs _
          Also, this helps fulshing the disk buffers to the file
        DoEvents
      Loop

  Next                                                            ' Until BytesDone = SourceBytes
  Close 2                                                         ' Close the source file

  If Compress = 1 Then
    On Error Resume Next
    Set Huffman = New clsHuffman
    frmProgress.lblCurrentFile = GetFileName(FileName)
    frmProgress.lblStatus = "DeCompressing..."
    Huffman.DecodeFile FNameNoExt & ".tmp", DestinationPath & FNameNoExt & Extension
    If FileLen(FNameNoExt & ".tmp") > 0 Then Kill FNameNoExt & ".tmp" ' Delete tmp file if size is larger than 0
  End If

  'When the code reaches this point, everything went OK.
  'Acknowledge the number of segments, assign the value '0' to the function and exit
  NumOfSegments = SegmentNumber
  JoinFile = 0
  frmProgress.picProgress.Cls
  frmProgress.Hide
  Exit Function

ErrorHandler:

    'This is entered only when an error occures
    Select Case ErrorCode
        Case 0 'Unknown error
            Reset   'Close any open files
            JoinFile = 4   'Assign error code 4 to the function
        Case Else 'Assign error code value to the function (1 to 3)
            JoinFile = ErrorCode
    End Select
    
    Exit Function

End Function
