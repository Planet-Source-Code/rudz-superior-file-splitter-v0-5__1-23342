Attribute VB_Name = "modProgressDeluxe"
' Code example by Rudy Alex Kohn
' Use as you like, but plz credit me for it =)
' rudyalexkohn@hotmail.com
' v1.01
' - Added option to change to line color (still 100% compatible)
' NOTE!!.. AutoRedraw MUST be set to TRUE!!
Option Explicit

Sub DrawPercent(picBox As PictureBox, lPercent As Integer, Optional ForeColor As Long = vbBlack, Optional LineColor As Long = vbBlue, Optional BackColor As Long = &H8000000F)
    picBox.Scale (0, 0)-(100, 100)                                          ' Set scale
    With picBox
        .BackColor = BackColor                                              ' Set Background color
        .ForeColor = ForeColor                                              ' Sets forecolor (%)
        .Cls                                                                ' Clear
        picBox.Line (0, 0)-(lPercent, 100), LineColor, BF                   ' The line update
        .CurrentX = (.ScaleWidth - .TextWidth(CStr(lPercent & " %"))) / 2   ' X
        .CurrentY = (.ScaleHeight - .TextHeight(CStr(lPercent & " %"))) / 2 ' Y
        picBox.Print CStr(lPercent) & " %"                                  ' Print xx%
    End With
End Sub
