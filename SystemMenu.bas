Attribute VB_Name = "SytemMenu"
'thank you to VbWebExample for the original example, and to Aerodynamica Software's enhancements

Public ProcOld As Long

'catch messages and call windows procedures
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'menu apis
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

'api constants
Public Const WM_SYSCOMMAND = &H112
Public Const GWL_WNDPROC = (-4)

'add new consts for new items
Public Const IDM_ITEM1 As Long = 0
Public Const IDM_ITEM2 As Long = 1
Public Const IDM_ABOUT As Long = 2

Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'do not debug this procedure, it will crash VB
    Select Case iMsg
        Case WM_SYSCOMMAND
            Select Case wParam
                Case IDM_ABOUT
                    message = "Autosave v1.2"
                    message = message & vbCrLf & vbCrLf
                    message = message & "Programmed By:" & vbCrLf
                    message = message & "Daniel Vandersluis" & vbCrLf
                    message = message & "iNFiNiTi Studios" & vbCrLf & vbCrLf
                    message = message & "2000"
                    MsgBox message, vbInformation, "About..."
                    Call SetWindowLong(hWnd, GWL_WNDPROC, ProcOld)
                    ' Including this line prevents VB from crashing if execution is
                    ' stopped, however, it only lets the about menu item be used once
                    '
                    ' What I'd like to do is find out when the user clicks on the
                    ' system menu, then enable the sysmenu handling
                    Exit Function
                
                Case IDM_ITEM1
                    MsgBox "Item 1!", vbInformation, "Item 1"
                    Exit Function
                    
                Case IDM_ITEM2
                    MsgBox "Item 2!", vbInformation, "Item 2"
                    Exit Function
            End Select
    End Select
    
    'pass all messages on to VB and then return the value to windows
    WindowProc = CallWindowProc(ProcOld, hWnd, iMsg, wParam, lParam)
End Function

