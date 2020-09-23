Attribute VB_Name = "modMain"
Option Explicit
'Type declaration for ChunkSize variable
Type ChunkSize
    S12000  As String * 12000
    S6000   As String * 6000
    S3000   As String * 3000
    S1500   As String * 1500
    S500    As String * 500
    S100    As String * 100
    S25     As String * 25
    S5      As String * 5
    S1      As String * 1
End Type
'Declare the variable Bytes as of ChunkSize type
Public Bytes As ChunkSize


Sub Main()
    CoolBorder frmSplit.hWnd

    '   Load forms to memory
    Load frmSplit
    Load frmProgress
'    Load frmDialog
    
    '   Show main form
    frmSplit.Show
End Sub

Sub Write000(FileName As String, Path As String, NoOfFiles As String, Extension As String, Compressed As String)
    EnCrypt "RAK", NoOfFiles, "Number of files", Path, FileName
    EnCrypt "Rln", Compressed, "Compressed", Path, FileName
    EnCrypt "RFS", Extension, "Extension", Path, FileName
End Sub

Sub Progress(Porcent As Integer)
    ' This is used to show progress when handling files
    DrawPercent frmProgress.picProgress, Porcent, , vbWhite
    DoEvents
End Sub
