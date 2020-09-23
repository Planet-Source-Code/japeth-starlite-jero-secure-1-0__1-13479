Attribute VB_Name = "HideModule"
Public Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" _
    Alias "FindWindowExA" (ByVal hWnd1 As Long, _
    ByVal hWnd2 As Long, ByVal lpsz1 As String, _
    ByVal lpsz2 As String) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Public Const WM_COMMAND = &H111
Public Const MIN_ALL = 419
Public Const MIN_ALL_UNDO = 416
Public Const LB_SETHORIZONTALEXTENT = &H194

Public Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Public Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Public Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Public Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINT
    X As Long
    Y As Long
End Type

'Other APIs
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Minimize all the windows on the desktop (and optionally restore them)
' This has the same effect as pressing the Windows+M key combination

Sub MinWindows(Optional Restore As Boolean)
    Dim hWnd As Long
    ' get the handle of the taskbar
    hWnd = FindWindow("Shell_TrayWnd", vbNullString)
    ' Minimize or restore all windows
    If Restore Then
        SendMessage hWnd, WM_COMMAND, MIN_ALL_UNDO, ByVal 0&
    Else
        SendMessage hWnd, WM_COMMAND, MIN_ALL, ByVal 0&
    End If
End Sub


Public Function WHide()
'System Bar
Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 0
'Ctrl-Alt-Delete
Dim ret As Integer
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, vbNullString, 0)
'Icons
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 0
'Minimize All Windows
MinWindows False
'Constrict Mouse
SetCursorPos 400, 300 'Center Screen
End Function

Public Function WShow()
'System Bar
Dim Handle As Long
Handle& = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow Handle&, 1
'Ctrl-Alt-Delete
Dim ret As Integer
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, vbNullString, 0)
'Icons
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 5
'Normalize All Windows
MinWindows True
'Release Mouse
ClipCursor ByVal 0&
End Function

Public Function DisButtons()
Dim ret As Integer
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, vbNullString, 0)
End Function

