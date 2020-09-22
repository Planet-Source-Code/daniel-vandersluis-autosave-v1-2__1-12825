Attribute VB_Name = "ListWindows"
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal wIndx As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal insaft As Long, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal flgs As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDNEXT = 2
Global Const GWL_STYLE = (-16)
Global Const WS_VISIBLE = &H10000000
Global Const WS_BORDER = &H800000
Global Const IsTask = WS_VISIBLE Or WS_BORDER

Function TaskWindow(hwCurr As Long) As Long
    Dim lngStyle As Long
    lngStyle = GetWindowLong(hwCurr, GWL_STYLE)
    If (lngStyle And IsTask) = IsTask Then TaskWindow = True
End Function
