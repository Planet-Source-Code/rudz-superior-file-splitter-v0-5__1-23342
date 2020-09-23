VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSplit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3270
   ClientLeft      =   2385
   ClientTop       =   2085
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin TabDlg.SSTab SS 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      _Version        =   327681
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "          &Split          "
      TabPicture(0)   =   "frmSplit.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkCompress"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkSplitCRC"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "          &Join          "
      TabPicture(1)   =   "frmSplit.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdJoin"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkJoinCRC"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "          &About          "
      TabPicture(2)   =   "frmSplit.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label8"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label9"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label10"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label12"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Command4"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      Begin VB.CheckBox chkSplitCRC 
         Caption         =   "Check CRC"
         Enabled         =   0   'False
         Height          =   240
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkJoinCRC 
         Caption         =   "Check CRC"
         Enabled         =   0   'False
         Height          =   240
         Left            =   -74760
         TabIndex        =   26
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Help"
         Height          =   360
         Left            =   -70320
         TabIndex        =   24
         Top             =   2040
         Width           =   795
      End
      Begin VB.CommandButton cmdJoin 
         Caption         =   "Join"
         Height          =   375
         Left            =   -71520
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Join"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   10
         Top             =   480
         Width           =   5175
         Begin VB.CommandButton cmdBrowse 
            Caption         =   ".."
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   4680
            TabIndex        =   14
            Top             =   360
            Width           =   315
         End
         Begin VB.TextBox txtFile 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   13
            Top             =   360
            Width           =   3375
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   ".."
            Height          =   285
            Index           =   1
            Left            =   4680
            TabIndex        =   12
            Top             =   840
            Width           =   315
         End
         Begin VB.TextBox txtFile 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   11
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "000 File :"
            Height          =   240
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label3 
            Caption         =   "Destination Folder :"
            Height          =   480
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1005
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Split"
         Height          =   1335
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   5175
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1200
            TabIndex        =   6
            Top             =   360
            Width           =   3375
         End
         Begin VB.CommandButton Command2 
            Caption         =   ".."
            Height          =   285
            Left            =   4680
            TabIndex        =   5
            Top             =   360
            Width           =   315
         End
         Begin VB.ComboBox cmbSize 
            Height          =   360
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "File name :"
            Height          =   240
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Segment size :"
            Height          =   240
            Left            =   120
            TabIndex        =   7
            Top             =   750
            Width           =   1080
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Split"
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chkCompress 
         Caption         =   "&Compress"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   2160
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   ".000 Encryption original by Rami Rasanen"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -74880
         TabIndex        =   25
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label10 
         Caption         =   "Join algorythm (c) 2001, Rudy Alex Kohn - Freeware"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -74880
         TabIndex        =   23
         Top             =   1800
         Width           =   3120
      End
      Begin VB.Label Label9 
         Caption         =   "Original Huffman compression algorythm (c) 2000, Fredrik Qvarfort"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -74880
         TabIndex        =   22
         Top             =   1560
         Width           =   5445
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "rudyalexkohn@hotmail.com"
         Height          =   240
         Left            =   -73140
         TabIndex        =   21
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "(C) 2001 By Rudy Alex Kohn"
         Height          =   240
         Left            =   -73125
         TabIndex        =   20
         Top             =   720
         Width           =   2025
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Superior File Splitter v0.5"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73560
         TabIndex        =   19
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Original split algorythm (c) Kamal A. Mehdi / Consumer Software Solutions"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -74880
         TabIndex        =   18
         Top             =   1320
         Width           =   4590
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Huffman As clsHuffman
Private sCurrentJoinFile As String
Private sOriginalJoinFile As String

Private Sub cmbSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call SendMessage(cmbSize.hWnd, CB_SHOWDROPDOWN, True, 0&)
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
    
    Select Case Index
    Case 0
        Call CoolBorder(frmDialog.hWnd)
        frmDialog.Show
        Exit Sub
    Case 1
        With CommonDialog1
            .DialogTitle = "Select .000 File"
            .Filter = "Info File (*.000)|*.000"
            .FileName = vbNullString
            .ShowOpen
            If LenB(.FileName) <> 0 Then txtFile(1) = .FileName
        End With
    End Select

End Sub

Private Sub cmdClose_Click()
    Unload Me
'    Unload frmDialog
    Unload frmProgress
    End
End Sub

Private Sub cmdJoin_Click()

    Dim sDestinationPath As String
    Dim sTmp As String

    sDestinationPath = txtFile(0).Text
    If LenB(sDestinationPath) = 0 Then
        sDestinationPath = App.Path & "\"        ' Set app.path if nothing is there
    ElseIf Not Right$(sDestinationPath, 1) = "\" Then
        sDestinationPath = sDestinationPath & "\"         ' Bug fix. just to be sure :)
    End If
    Dim sFilename As String
    sFilename = txtFile(1)
    Select Case LenB(sFilename)
    Case 0                                                         ' Field is empty
        MsgBox "Info File not selected.", 16, sTitle
        Exit Sub
    Case Else
        Dim InfoFile As Boolean
'        Dim bComp As Boolean
        InfoFile = True
        If Right$(sFilename, 3) = "000" Then         ' File is not a valid info file
            Open sFilename For Input As #1
                Line Input #1, sTmp
            Close #1
            If sTmp = "[File Splitter]" Then
                InfoFile = True
            End If
        End If
'                bComp = False
'            ElseIf Left(sTmp, 2) = "HE" Then
'                bComp = True
'            Else
'                InfoFile = False
'            End If
 '       End If
        sCurrentJoinFile = sFilename
        sOriginalJoinFile = sFilename
        If InfoFile = False Then
            MsgBox GetFileName(sFilename) & " is not a valid Info File, please select using the dialog!", 16, sTitle
            Exit Sub
'        ElseIf (bComp) Then
'            Open sFilename For Binary Access Read As #1
'                Get #1, , sTmp ' = input(EOF(1), 1)
'            Close #1
'            MsgBox sTmp
'            Set Huffman = New clsHuffman
'            sTmp = Huffman.DecodeString(sTmp)
''            sTmp = DeCompressString(sTmp)
'            MsgBox sTmp
'            Open GetFileName(sFilename) & ".tmp" For Output As #1
'                Print #1, sTmp
'            Close #1
'            sCurrentJoinFile = GetFileName(sFilename) & ".tmp"
            ' Do Unpacking of file here!!
'        Else
'            sCurrentJoinFile = sOriginalJoinFile
        End If
            
    End Select


    Dim NumberOfFiles As String
    Dim Compressed As String
    Dim Extension As String
'    Dim CRC() As String

    Dim i As Integer
    Dim sPath As String

    sPath = GetFilePath(sFilename)
    sFilename = GetFileName(sFilename)

    NumberOfFiles = DeCrypt("RAK", "Number of files", ".\", sFilename)
'    MsgBox NumberOfFiles
    Compressed = DeCrypt("Rln", "Compressed", ".\", sFilename)
'    MsgBox Compressed
    Extension = DeCrypt("RFS", "Extension", ".\", sFilename)
'    MsgBox Extension

    With frmProgress
        .Show
        .lblCurrentFile = GetFileNoExtension(sFilename) & ".001"
        .lblStatus = "Reading"
        .Refresh
    End With
    i = JoinFile(GetFileNoExtension(sFilename) & Extension, Extension, Val(Compressed), Val(NumberOfFiles), sPath)
    If i = 0 Then
        MsgBox "OK!"
    Else
        MsgBox "AARGH!"
    End If
    ' Call CRC Check if user has selected it

    ' Call JoinFile -- HERE --
End Sub

Private Sub Command1_Click()

    ' Firstly, check if values are correct
    Select Case ""
    Case Text1
        MsgBox "No input file selected.", vbCritical, Me.Caption
        Text1.SetFocus
        Exit Sub
    Case cmbSize.Text
        MsgBox "No segment size selected.", vbCritical, Me.Caption
        cmbSize.Text = "1.00 Mb"
        Exit Sub
    End Select
    
    
    Dim x As Integer
    Dim Segments As Integer
    Dim SourceFile As String
    Dim SegmentSize As Long
    
    With cmbSize
    Select Case .Text
    Case "1.00 Mb"
        SegmentSize = 1000
    Case "2.88 Mb"
        SegmentSize = 2880
    Case "1.44 Mb"
        SegmentSize = 1440
    Case "5.00 Mb"
        SegmentSize = 5000
    Case "100 Kb"
        SegmentSize = 100
    Case "250 Kb"
        SegmentSize = 250
    Case "500 Kb"
        SegmentSize = 500
    Case "720 Kb"
        SegmentSize = 720
    Case "7.50 Mb"
        SegmentSize = 7500
    Case "10.0 Mb"
        SegmentSize = 10000
    Case "25.0 Mb"
        SegmentSize = 25000
    Case Else
        SegmentSize = 1000
    End Select
    End With
    SegmentSize = SegmentSize * 1024
    SourceFile = Text1

    frmProgress.lblStatus = "Retrieving Info..."
    frmProgress.Show , frmSplit
    frmProgress.Refresh

    'Call the function
    x = SplitFile(SourceFile, SegmentSize, chkCompress.Value, Segments)

    'Inform the user about the call success or failure
    Dim sMsg As String, nMsgDis As Long
    Select Case x
    Case 0
        frmProgress.picProgress.Cls
        sMsg = "The process completed successfully." & vbCr & "The file was split to " & Segments & " segments."
        sMsgDis = 64
    Case 1
        sMsg = "File does not exists."
        sMsgDis = 16
    Case 2
        sMsg = "Invalid segment size."
        sMsgDis = 16
    Case 3
        sMsg = "Unable to create more than 999 segments." & vbCr & "Please raise segment size and try again."
        sMsgDis = 16
    Case 5
        sMsg = "Resulting file is smaller than requested segment size." & vbCr & "Please try another size or don't split the file at all :)"
        sMsgDis = 64
    Case 4
        sMsg = "Unknown error!!!!!"
        sMsgDis = 16
    End Select
    Call MsgBox(sMsg, sMsgDis, sTitle)
    frmProgress.picProgress.Cls
    frmProgress.Hide
End Sub

Private Sub Command2_Click()

    'Initialize the common dialog control and show it
    With CommonDialog1
        .DialogTitle = "Select file to split"
        .Filter = "All Files (*.*)|*.*"
        .FileName = vbNullString
        .ShowOpen
        If LenB(.FileName) <> 0 Then Text1 = .FileName
    End With
End Sub

Private Sub Form_Load()
    With cmbSize
        .AddItem "100 Kb"
        .AddItem "250 Kb"
        .AddItem "500 Kb"
        .AddItem "720 Kb"
        .AddItem "1.00 Mb"
        .AddItem "1.44 Mb"
        .AddItem "2.88 Mb"
        .AddItem "5.00 Mb"
        .AddItem "7.50 Mb"
        .AddItem "10.0 Mb"
        .AddItem "25.0 Mb"
        .Text = "1.00 Mb"
    End With
    Me.Caption = sTitle
    lblStatus = "Status : Idle..."
    SS.Tab = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If LenB(Text1) = 0 Then Command2_Click Else Command1_Click
End Sub
