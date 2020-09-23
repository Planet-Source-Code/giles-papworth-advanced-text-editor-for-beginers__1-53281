Attribute VB_Name = "modmain"
Option Explicit
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const LB_ITEMFROMPOINT = &H1A9

Public Declare Function Setwindowpos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Function SetWinPos(iPos As Integer, lHWnd As Long) As Boolean
    Dim lwinpos As Long
    iPos = 1
    Select Case iPos
        Case 1
            lwinpos = HWND_TOPMOST
        End Select
    If Setwindowpos(lHWnd, lwinpos, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE) Then
        SetWinPos = True
    End If
End Function


