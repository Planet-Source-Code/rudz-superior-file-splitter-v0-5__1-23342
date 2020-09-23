Attribute VB_Name = "modSplit"
Option Explicit


Private Huffman As clsHuffman
Public Const sTitle As String = "Superior File-Splitter v0.5"


Function SplitFile(FileName As String, SegmentSize As Long, Compress As Integer, Optional NumOfSegments As Integer) As Integer

    On Error GoTo ErrorHandler

    
    'Make sure the file exists
    Dim ErrorCode As Integer
    If LenB(FileName) = 0 Or LenB(Dir(FileName)) = 0 Then
        ErrorCode = 1
        GoTo ErrorHandler
    ElseIf SegmentSize = 0 Then 'Ensure that the segment size is valid
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
    Dim sOrgFileExt         As String               '   Store original file extension

'    Dim i As Integer


    FName = GetFileName(FileName)                   '   Extract the file name
    FPath = GetFilePath(FileName)                   '   Extract the path name

    FNameNoExt = GetFileNoExtension(FName)          '   File name without extension
    sOrgFileExt = GetFileExtension(FName)

    If Compress = 1 Then
        With frmProgress
        Set Huffman = New clsHuffman
        .lblCurrentFile = GetFileName(FileName)
        .lblStatus = "Compressing..."
        Huffman.EncodeFile FileName, FNameNoExt & ".tmp"
        Dim sTmp As String
        sTmp = FNameNoExt & ".tmp"
        .lblStatus = "Checking Size..."
        Select Case FileLen(sTmp)
        Case Is < SegmentSize ' If filesize is lower then segmentsize, then ask if use want's to keep it anyway
            If MsgBox("Selected file was successfuly compressed to " & FileLen(sTmp) & "bytes." & vbCr & "Do you wish to keep this file?", vbYesNo, sTitle) = vbYes Then
                Write000 FNameNoExt & ".000", FPath, "1", sOrgFileExt, "1"
                .lblStatus = "Copying..."
                FileCopy sTmp, FNameNoExt & ".001"
                .lblStatus = "Deleting..."
                Kill sTmp
                NumOfSegments = 1
            Else
                .lblStatus = "Deleting..."
                Kill sTmp
            End If
            .lblStatus = "Idle..."
            Exit Function
        Case Is < FileLen(FileName)
            FileName = sTmp
        Case Is > FileLen(FileName)
            Kill sTmp
            Compress = 0
        End Select
        End With
    End If

    'Get total number or bytes in the source file
    SourceBytes = FileLen(FileName)

    'Ensure that the resultant file segments will not exceed 999 segments
    'because otherwise we will have incorrect file extensions
    If SourceBytes / SegmentSize >= 1000 Then
        ErrorCode = 3
        GoTo ErrorHandler
    End If

    'Open the source file for binary read
    Open FileName For Binary Access Read As #1 Len = 1

    SegmentNumber = 0
    Do
        'Increase the number of segments counter by 1
        SegmentNumber = SegmentNumber + 1

        'Compose the file name of the new file to be created (file segment)
        Select Case SegmentNumber
            Case Is < 10
                DestinationFile = FPath & FNameNoExt & ".00" & CStr(SegmentNumber)
            Case 10 To 99
                DestinationFile = FPath & FNameNoExt & ".0" & CStr(SegmentNumber)
            Case 100 To 999
                DestinationFile = FPath & FNameNoExt & "." & CStr(SegmentNumber)
        End Select



        'Create the new file segment and open it for binary write
        Open DestinationFile For Binary Access Write As #2 Len = 1

        'Check whether the remaining bytes to process in the source file are
        'less than Segment bytes
        Select Case SourceBytes - BytesDone
        Case Is < SegmentSize
            RemainingBytes = SourceBytes - BytesDone
        Case Else
            RemainingBytes = SegmentSize
        End Select

        frmProgress.lblStatus = "Writing...."
        frmProgress.lblCurrentFile = GetFileName(DestinationFile)
       'Read bytes from the source file and write them to the destination file (the current segment file)
       'Depending on the remaining bytes to read and write, the routine below will read the largest possible
       'chunk of data
        With Bytes
        Do
            Select Case RemainingBytes
                Case Is >= 12000
                    '   Read 12000 bytes of data from the source file
                    Get #1, , .S12000
                    Put #2, , .S12000
                    '   Decrease the number of remaining bytes by 12000
                    RemainingBytes = RemainingBytes - 12000
                    '   Update the bytes done counter
                    BytesDone = BytesDone + 12000
                    '   Yield to windows and other processes to do their jobs
                    '   Also, this helps fulshing the disk buffers to the file
                    DoEvents
                Case 6000 To 11999
                    Get #1, , .S6000
                    Put #2, , .S6000
                    RemainingBytes = RemainingBytes - 6000
                    BytesDone = BytesDone + 6000
                    DoEvents
                Case 3000 To 5999
                    Get #1, , .S3000
                    Put #2, , .S3000
                    RemainingBytes = RemainingBytes - 3000
                    BytesDone = BytesDone + 3000
                    DoEvents
                Case 1500 To 2999
                    Get #1, , .S1500
                    Put #2, , .S1500
                    RemainingBytes = RemainingBytes - 1500
                    BytesDone = BytesDone + 1500
                    DoEvents
                Case 500 To 1499
                    Get #1, , .S500
                    Put #2, , .S500
                    RemainingBytes = RemainingBytes - 500
                    BytesDone = BytesDone + 500
                    DoEvents
                Case 100 To 499
                    Get #1, , .S100
                    Put #2, , .S100
                    RemainingBytes = RemainingBytes - 100
                    BytesDone = BytesDone + 100
                    DoEvents
                Case 25 To 99
                    Get #1, , .S25
                    Put #2, , .S25
                    RemainingBytes = RemainingBytes - 25
                    BytesDone = BytesDone + 25
                    DoEvents
                Case 5 To 24
                    Get #1, , .S5
                    Put #2, , .S5
                    RemainingBytes = RemainingBytes - 5
                    BytesDone = BytesDone + 5
                    DoEvents
                Case 1 To 4
                    Get #1, , .S1
                    Put #2, , .S1
                    RemainingBytes = RemainingBytes - 1
                    BytesDone = BytesDone + 1
                    DoEvents
                Case 0
                    '   When the loop enters here, the segment bytes are completed.
                    '   Close the segment file and exit the loop
                    Close 2
                    DoEvents
                    Exit Do
            End Select
            'Update the percent control on the form
            DrawPercent frmProgress.picProgress, Int((BytesDone / SourceBytes) * 100), , vbWhite
            'Refresh the form and yield to windows
            DoEvents
        Loop
        End With
    Loop Until BytesDone = SourceBytes
    'Close the source file
    Close 1
    frmProgress.lblStatus = "Writing..."
    sTmp = FNameNoExt & ".000"
    frmProgress.lblCurrentFile = sTmp

    Write000 sTmp, FPath, str$(SegmentNumber), sOrgFileExt, str$(Compress)

    'When the code reaches this point, everything went OK.
    'Acknowledge the number of segments, assign the value '0' to the function and exit
    NumOfSegments = SegmentNumber
    On Error Resume Next
    sTmp = GetFileNoExtension(FileName) & ".tmp"
    If FileLen(sTmp) > 0 Then Kill sTmp
    SplitFile = 0
    frmProgress.lblStatus = "Idle..."
    frmProgress.lblCurrentFile = vbNullString
    Exit Function

ErrorHandler:
    '   This is entered only when an error occures
    Select Case ErrorCode
        Case 0                      '   Unknown error
            Reset                   '   Close any open files
            SplitFile = 4           '   Assign error code 4 to the function
        Case Else                   '   Assign error code value to the function (1 to 3)
            SplitFile = ErrorCode
    End Select
    Exit Function
End Function

Function CompressTry(StringToCompress As String, Optional NumberOfBytesToTest As Long = 1024) As Integer
' Returns 1 if it pays to compress, 0 if it doesn't

    Dim sUnCompressed As String
    Dim sCompressed As String
    Dim nUnCompressed As Long
    Dim nCompressed As Long

    ' First, get the lenght of the original bytes to test
    sUnCompressed = Mid$(StringToCompress, 1, NumberOfBytesToTest)
    nUnCompressed = Len(sUnCompressed)


    ' Then, compress exactly the same bytes
    sCompressed = Huffman.EncodeString(Mid$(StringToCompress, 1, NumberOfBytesToTest))
    nCompressed = Len(sCompressed)

    Select Case nUnCompressed
    Case Is > nCompressed
        CompressTry = 1
    Case Else
        CompressTry = 0
    End Select
    
    nCompressed = vbNull
    sCompressed = vbNullString
    sUnCompressed = vbNullString
    nUnCompressed = vbNull
End Function

