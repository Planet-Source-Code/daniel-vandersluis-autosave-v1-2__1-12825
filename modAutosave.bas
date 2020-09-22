Attribute VB_Name = "Various"
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal insaft As Long, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal flgs As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

' Constants for Screen Metrics
Const SM_CXFULLSCREEN = 16
Const SM_CYFULLSCREEN = 17

' Constants for Window Position
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const FLAGS = SWP_NOMOVE & SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Global LogPath As String
Public Sub CenterForm(frm As Form)
    ' Centers a form on the screen taking the taskbar into consideration.
    ' Basically what this does is take the full screen size excluding the taskbar
    ' and centers your form accordingly.
    
    Dim Left As Long, Top As Long
    
    Left = (Screen.TwipsPerPixelX * (GetSystemMetrics(SM_CXFULLSCREEN) / 2)) - _
        (frm.Width / 2)
    Top = (Screen.TwipsPerPixelY * (GetSystemMetrics(SM_CYFULLSCREEN) / 2)) - _
        (frm.Height / 2)
    frm.Move Left, Top
End Sub

Public Sub AlwaysOnTop(f As Form, pos As Boolean)

    If pos = True Then
        Call SetWindowPos(frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
        ' SetWindowPos takes a hWnd of a window, and moves it to a specific
        ' screen location (specified by the 3rd through 6th parameters) and
        ' changes the Z-Order as requested.
    Else
        SetWindowPos frmMain.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
    End If
    
End Sub

