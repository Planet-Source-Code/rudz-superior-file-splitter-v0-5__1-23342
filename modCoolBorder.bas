Attribute VB_Name = "modCoolBorder"
' Wrote by ?
' Adapted by Rudy Alex Kohn
Option Explicit

'> CoolBorder
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
'< CoolBorder

Sub CoolBorder(lHwnd As Long)
' adds cool look to form >
' Call CoolBorder(Me.hWnd)
    Dim lRet As Long
    lRet = GetWindowLong(lHwnd, GWL_EXSTYLE)
    lRet = lRet Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    SetWindowLong lHwnd, GWL_EXSTYLE, lRet
    SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub
