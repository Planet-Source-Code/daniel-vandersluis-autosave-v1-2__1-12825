VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoSave"
   ClientHeight    =   7365
   ClientLeft      =   150
   ClientTop       =   225
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   100
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   6570
   Begin VB.ComboBox cboTime 
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   4860
      Width           =   975
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "&Empty Log"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "Save a &Logfile"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   6420
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   5880
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer tmrKey 
      Left            =   0
      Top             =   1560
   End
   Begin VB.CheckBox chkVisible 
      Caption         =   "Show only &Visible Windows"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   6090
      Width           =   2415
   End
   Begin VB.TextBox txtRemaining 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0"
      Top             =   5520
      Width           =   975
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   10
      Left            =   0
      Top             =   1080
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set &Timer"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      Top             =   4860
      Width           =   1215
   End
   Begin VB.CheckBox chkCurrent 
      Caption         =   "Autosave in &Current Window"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CheckBox chkSaveAs 
      Caption         =   "Activate &Save As Dialog Box"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   5400
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Timer tmrLoop 
      Interval        =   1000
      Left            =   0
      Top             =   600
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   120
   End
   Begin VB.OptionButton optOff 
      Caption         =   "Off"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   4680
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optOn 
      Caption         =   "On"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   4200
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   3765
      ItemData        =   "Main.frx":014A
      Left            =   360
      List            =   "Main.frx":014C
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh List"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   4200
      WhatsThisHelpID =   1
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Seconds Remaining:"
      Height          =   195
      Left            =   3840
      TabIndex        =   16
      Top             =   5520
      Width           =   1440
   End
   Begin VB.Label lblTimerInterval 
      AutoSize        =   -1  'True
      Caption         =   "Timer Interval:"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   4920
      Width           =   1065
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'AUTOSAVE PROGRAM v1.2
'
'Written by:    Daniel Vandersluis
'               E-Mail: dvandersluis@icqmail.com
'               ICQ:    57887382
'Version:       1.2
'Finished:      05-Nov-2000
'
'Description:   A lot of programs do not automatically save your work
'               and therefore, if something like a crash occurs, all of
'               your ammendments to your file since your last save are lost.
'               This program automatically saves your file for you, saving
'               potentially losing your work.
'
'Updates:       Version 1.2
'                   * Added Enabled, Disabled, and Set Timer to the systray menu
'                   * Fixed problem of Timer restrictions
'                   * Added menu items to system menu
'                   * Added option to view only visible windows
'                   * Removed limit on proc title size
'                   * Fixed bugs encountered with VB6
'                   * Added DX7 DirectInput support for System Hotkeys
'                   * Added the progress bar
'                   * Added the option to save a log
'                   * Created a new icon
'                   * Allowed interval to be in seconds or minutes
'                   * Added a helpfile and popup menus to access it with contextIDs
'
'Programming:   Uses many system API calls, but everything is documented, so
'               don't worry. I tried to comment as much as I could. I'd say that
'               this is pretty Advanced if you would like a difficulty rating.
'
'Distribution:  I don't mind if you use part or all of this code within your
'               own programs, however, I ask that you do not distribute it as
'               your own program and give credit where credit is deserved - Please
'               either link to my website or to my PSC submission page if you use this
'               code.
'
'Feedback:      I'd appreciate feedback, comments, and/or suggestions you might have.
'               Either log them through PSC, email them to me, or ICQ me and I'll
'               respond as soon as I can.
'
'Thanks to:     * Pause Break [mofd4u@yahoo.com] for supplying PSC with his
'                 Registry module
'               * Nick Smith aka ImN0thing for the icon in systray code
'               * Bryan Stafford of New Vision Software® - newvision@imt.net - for
'                 the menu columns code
'               * The Ian Ippolito, creator of Planet Source Code, and Exhedra
'                 Solutions, Inc. for their great system
'
'By the way:    Check out http://www.angelfire.com/on3/infiniti/index.html
'               This is my homepage and it contains all my programs in C/C++ and
'               VB, as well as Descent 3 Levels, and much more, all with open
'               source code!
'============================================================================

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lParam As Long) As Long

Public TimerSet As Boolean
Public application As String
Public listNow As Integer
Public listNum As Integer
Dim hnd, X
Dim done As Boolean

'Initialize DirectInput - needs dx7vb.dll (included in zip file)
Dim dx As New DirectX7  'the directX object.
Dim di As DirectInput   'the directInput object.
Dim diDEV As DirectInputDevice  'the sub device of DirectInput.
Dim diState As DIKEYBOARDSTATE  'the key states.
Dim iKeyCounter As Integer
Dim aKeys(255) As String    'key names

Private Sub cboTime_Click()

    On Error Resume Next
    If UCase(cboTime.Text) = "SECONDS" Then
        txtInterval.Text = Format(txtInterval.Text * 60, "#")
        retVal = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "TimeMeasurement", "Seconds")
    ElseIf UCase(cboTime.Text) = "MINUTES" Then
        txtInterval.Text = Format(txtInterval.Text / 60, "#.##")
        retVal = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "TimeMeasurement", "Minutes")
    End If

End Sub

Private Sub chkCurrent_Click()

    retVal = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "Current", chkCurrent.Value)
    'UpdateKey creates/updates a registry entry
    
End Sub

Private Sub chkCurrent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 7
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100
    
End Sub

Private Sub chkSave_Click()

    retVal = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "SaveLog", chkSave.Value)
    
End Sub

Private Sub chkSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 9
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100
    
End Sub

Private Sub chkSaveAs_Click()
    
    retVal = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "SaveAs", chkSaveAs.Value)
    'See chkCurrent_Click()
    
End Sub

Private Sub chkSaveAs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 7
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100
    
End Sub

Private Sub chkVisible_Click()

    retVal = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "Visible", chkVisible.Value)
     ' For comments on window enumeration, see Form_Activate()
    
    Erase ProcInfo
    NumProcs = 0
    
    List1.Clear
    EnumWindows AddressOf EnumProc, 0

    tmrWait.Enabled = True
    
End Sub

Private Sub chkVisible_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 8
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub

Private Sub cmdKill_Click()

    Dim fPath
    fPath = App.Path
    If Right(fPath, 1) <> "\" Then fPath = fPath + "\"
    fPath = fPath + "autosave.log"
    On Error Resume Next
    Kill fPath

End Sub

Private Sub cmdKill_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.HelpContextID = 9
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100
    
End Sub

Private Sub cmdSet_Click()
    
    If txtInterval.Text > 0 Then
        On Error Resume Next
        If cboTime.Text = "Seconds" Then
            TimerSet = True
            txtRemaining = txtInterval.Text
            ProgressBar1.Max = txtInterval.Text
        ElseIf cboTime.Text = "Minutes" Then
            TimerSet = True
            ProgressBar1.Max = Format(txtInterval.Text * 60, "#")
            txtRemaining = Format(txtInterval.Text * 60, "#")
        End If
        ' Update Registry Key
            retVal = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "TimerInterval", txtRemaining)
        ProgressBar1.Value = 0
    End If

End Sub

Private Sub cmdSet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 5
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub

Private Sub Command1_Click()

    ' For comments on window enumeration, see Form_Activate()
    
    Erase ProcInfo
    NumProcs = 0
    
    List1.Clear
    EnumWindows AddressOf EnumProc, 0

    tmrWait.Enabled = True
    
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 4
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub

Private Sub Form_Activate()

    hnd = GetSystemMenu(frmMain.hWnd, False)
    X = AppendMenu(hnd, 0, IDM_ABOUT, "About...")
    
    ' Enumerating current windows into a list box:
    
    ' Erase any old information.
    Erase ProcInfo
    NumProcs = 0
    
    List1.Clear
    EnumWindows AddressOf EnumProc, 0

    ' Wait for the enumeration to finish.
    tmrWait.Enabled = True

End Sub
Private Sub Form_Load()

    Set di = dx.DirectInputCreate()
    'create the object, must be done before anything else
    If Err.Number <> 0 Then 'if err=0 then there are no errors.
        MsgBox "Error starting Direct Input, please make sure DirectX is installed", vbApplicationModal
        End
    End If
    Set diDEV = di.CreateDevice("GUID_SysKeyboard")
    'Create a keyboard object off the Input object
    diDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    'specify it as a normal keyboard, not mouse or joystick
    diDEV.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    ' set coop level. Defines how it interacts with other applications,
    ' whether it will share with other apps. DISCL_NONEXCLUSIVE means that
    ' it's multi-tasking friendly
    diDEV.Acquire   'aquire the keystates.
    tmrKey.Interval = 10    'sensitivity, in this case the repeat rate of the keyboard
    tmrKey.Enabled = True   'enable the timer, this has the key detecting code in it
    
    CenterForm Me
    
    done = False
    
    'Get Stored Settings from Registry
    Dim tInterval As String
    On Error Resume Next
    txtInterval = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "TimerInterval")
    If txtInterval = "" Then txtInterval = 0
    ProgressBar1.Max = txtInterval.Text
    chkSaveAs.Value = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "SaveAs")
    chkSave.Value = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "SaveLog")
    chkCurrent.Value = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "Current")
    chkVisible.Value = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "Visible")
    cboTime.Text = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\iNFiNiTi Studios\Autosave\", "TimeMeasurement")
    If UCase(Left(cboTime.Text, 5)) = "FALSE" Then cboTime.Text = "Seconds"
    If cboTime.Text = "Minutes" Then txtInterval = Format(txtInterval / 60, "#.##")
    txtRemaining = txtInterval.Text
    
    'Hide window from TaskList (CTRL+ALT+DEL List)
    OwnerhWnd = GetWindow(Me.hWnd, 4)   ' Specifies what part of the window to get
                                        ' Here OwnerhWnd is the part of the window
                                        ' for the TaskList
    retVal = ShowWindow(OwnerhWnd, 0)   ' Hides the window from the TaskList
    
    Dim temp
    
    'Create Systray icon
    NID.hWnd = Me.hWnd
    NID.cbSize = Len(NID)
    NID.uID = vbNull
    NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    NID.hIcon = Me.Icon
    NID.uCallbackMessage = WM_MOUSEMOVE
    SetTrayTip "Autosave"
    
    Shell_NotifyIcon NIM_ADD, NID

    TimerSet = False
    
    listNum = NumProcs
    
    ' This line is needed to be able to handle the system menu, however if you
    ' try to stop execution with this line included, VB will crash
    '
    ' In Windows, each form that is created has a default procedure which it uses
    ' to have events. This is why, unlike Macs, when you click the close button on
    ' a form, it closes, without adding any code.
    ' To control the system menu, we have to divert control from the
    ' default form procedure, WindowProc, to our procedure.
    ' When this happens, until it is reset, control over the system menu is passed
    ' to our procedure. And when the program is terminated properly, we return
    ' control to the default WindowProc. However, when we do not terminate
    ' properly, i.e pressing the stop button in VB, control is not returned to the
    ' system, and therefore VB crashes since there is nothing controlling it.
    ' Until I can find away around this, I am leaving this line out, but this is the
    ' line that diverts control to us.
    '
    ' ProcOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WindowProc)
  
    cboTime.AddItem "Seconds"
    cboTime.AddItem "Minutes"
  
    Me.Show 'show the form
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg As Long
Msg = X / Screen.TwipsPerPixelX
    Select Case Msg
                
        Case WM_LBUTTONDBLCLK
        ' if the left mouse button is double-clicked on the systray icon, the window
        ' is restored
        Me.WindowState = 0
        optOff.Value = True
        ShowWnd (Me.hWnd)
        SetForegroundWindow (Me.hWnd)
        
        Case WM_RBUTTONUP
        ' if the right mouse button is pressed on the systray icon, a menu pops up
        Dim pAPI As POINTAPI
        Dim PMParams As TPMPARAMS
        
        'get point at which the user clicked'
        GetCursorPos pAPI
        
        'create a new popup menu, and add entries'
        'and assign ids to each'
        Dim tmpPop&
        tmpPop& = CreatePopupMenu
        
        Dim ListCounter As Integer
        
        ' Insert all window titles from the list box into the first part of the
        ' popup menu
        
        'This is my old code that wasn't hierarchal:
            'For ListCounter = 0 To NumProcs - 1
            '    InsertMenu tmpPop%, 1 + ListCounter, MF_BYPOSITION, 200 + ListCounter, List1.List(ListCounter)
            'Next ListCounter
            
            'InsertMenu tmpPop%, NumProcs + 1, MF_SEPARATOR, 71, vbNullString
            'InsertMenu tmpPop%, NumProcs + 2, MF_BYPOSITION, 72, "Restore"
            'InsertMenu tmpPop%, NumProcs + 3, MF_BYPOSITION, 73, "Exit"
        
        Dim subMenu As Long
        
        subMenu = CreatePopupMenu
        
        Call InsertMenu(tmpPop&, 1&, MF_POPUP Or MF_STRING Or MF_BYPOSITION, subMenu&, "Procs")
        InsertMenu tmpPop&, 2&, MF_SEPARATOR, 71, vbNullString
        InsertMenu tmpPop&, 3&, MF_BYPOSITION, 72, "Enable" & vbTab & "F2"
        InsertMenu tmpPop&, 4&, MF_BYPOSITION, 73, "Disable" & vbTab & "F3"
        InsertMenu tmpPop&, 5&, MF_BYPOSITION, 74, "Set Timer..."
        InsertMenu tmpPop&, 6&, MF_SEPARATOR, 75, vbNullString
        InsertMenu tmpPop&, 7&, MF_BYPOSITION, 76, "Restore"
        InsertMenu tmpPop&, 8&, MF_BYPOSITION, 77, "Exit"
             
        For ListCounter = 0 To NumProcs - 1
            Call InsertMenu(subMenu, ListCounter + 0&, MF_STRING Or MF_BYPOSITION, 200& + ListCounter, List1.List(ListCounter))
            
            ' after every 20 menu items, create a new column in the menu:
            If (ListCounter Mod 20 = 0) And ListCounter <> 0 Then
                Call ModifyMenu(subMenu, ListCounter, MF_BYPOSITION Or MF_MENUBARBREAK, ListCounter, List1.List(ListCounter))
            End If
            If chkVisible = 1 And ListCounter = NumProcs - 2 Then Exit For
        Next ListCounter
        
        'this is a standard size required for
        'the popup menu to be displayed
        PMParams.cbSize = 20
        
        'display the popup menu, note the flag "TMP_RETURNCMD", it sets the value
        'of tmpReply% to the id of the menu item that was clicked
        tmpReply% = TrackPopupMenuEx(tmpPop&, TPM_LEFTALIGN Or TPM_LEFTBUTTON Or TPM_RETURNCMD, pAPI.X, pAPI.Y, Me.hWnd, PMParams)
        
        Select Case tmpReply%
            Case 200 To (200 + NumProcs - 1)    ' Select title in list box
                List1.Selected(tmpReply% - 200) = True
                
            Case 72 'On
                If optOn.Enabled = True Then optOn.Value = True
            
            Case 73 'Off
                optOff.Value = True
            
            Case 74 'Set Timer
                Dim retVal As Double
                
                retVal = 0
                rVal = InputBox("Please enter the timer interval:", "Set Timer", 10000)
                If rVal = "" Then rVal = txtInterval.Text
                txtInterval.Text = rVal
                On Error Resume Next
                'tmrLoop.Interval = txtInterval.Text
                TimerSet = True
                ' Update Registry Key
                retVal = UpdateKey(HKEY_LOCAL_MACHINE, "\Software\iNFiNiTi Studios\Autosave", "TimerInterval", txtInterval)
                txtRemaining = txtInterval.Text
                                           
            Case 76 'Restore
                optOff.Value = True
                ShowWnd Me.hWnd
                                
            Case 77 'Exit
                Dim NID As NOTIFYICONDATA
                
                NID.hWnd = Me.hWnd
                NID.cbSize = Len(NID)
                NID.uID = vbNull
                NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
                NID.hIcon = Me.Icon
                NID.uCallbackMessage = WM_MOUSEMOVE
                NID.szTip = "Right-Click to display Popupmenu"
                
                Shell_NotifyIcon NIM_DELETE, NID
                
                Call SetWindowLong(hWnd, GWL_WNDPROC, ProcOld)
                End
        End Select
    End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim NID As NOTIFYICONDATA
    
    ' Deletes the systray icon
    NID.hWnd = Me.hWnd
    NID.cbSize = Len(NID)
    NID.uID = vbNull
    NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    NID.hIcon = Me.Icon
    NID.uCallbackMessage = WM_MOUSEMOVE
    NID.szTip = "Right-Click to display Popupmenu"
    
    Shell_NotifyIcon NIM_DELETE, NID
    
    Call SetWindowLong(hWnd, GWL_WNDPROC, ProcOld)
    
End Sub
Private Sub Form_Resize()

'If the form is minimized, hide it from the taskbar
    If Me.WindowState = 1 Then
        HideWnd Me.hWnd
    Else ' If Restore is pressed
        ShowWnd Me.hWnd
    End If

End Sub
Private Sub Form_Terminate()
    
    Dim NID As NOTIFYICONDATA
    
    NID.hWnd = Me.hWnd
    NID.cbSize = Len(NID)
    NID.uID = vbNull
    NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    NID.hIcon = Me.Icon
    NID.uCallbackMessage = WM_MOUSEMOVE
    NID.szTip = "Right-Click to display Popupmenu"
    
    Shell_NotifyIcon NIM_DELETE, NID
    
    Call SetWindowLong(hWnd, GWL_WNDPROC, ProcOld)
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Dim NID As NOTIFYICONDATA
    
    NID.hWnd = Me.hWnd
    NID.cbSize = Len(NID)
    NID.uID = vbNull
    NID.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    NID.hIcon = Me.Icon
    NID.uCallbackMessage = WM_MOUSEMOVE
    NID.szTip = "Right-Click to display Popupmenu"
    
    Shell_NotifyIcon NIM_DELETE, NID
    
    Call SetWindowLong(hWnd, GWL_WNDPROC, ProcOld)
    diDEV.Unacquire
    
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 3
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub

Private Sub mnuHelp_Click()

    SendKeys "{F1}"

End Sub

Private Sub optOff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 10
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub

Private Sub optOn_Click()
    
    Dim lRetVal As Long
    Dim bRetVal As Boolean
    
    If chkCurrent.Value = False Then
        frmMain.WindowState = 1
        myHWnd = CLng(List2.List(List1.ListIndex))
        lRetVal = SetForegroundWindow(myHWnd)
        bRetVal = OpenIcon(myHWnd)
    End If
    
End Sub

Private Sub optOn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 10
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 6
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub

Private Sub tmrLoop_Timer()
    Dim lRetVal As Long
    Dim bRetVal As Boolean
    Dim myHWnd As Long
    Dim fName As String
    Dim fileNum
    
    If optOn.Value = True Then
    txtRemaining = txtRemaining - 1
    On Error Resume Next
    If txtRemaining > -1 Then ProgressBar1.Value = ProgressBar1.Value + 1
        If txtRemaining = -1 Then
            ProgressBar1.Value = ProgressBar1.Max
            fileNum = FreeFile
            Dim txt
            txt = Now
            txt = txt & " " & List1.List(List1.ListIndex)
            fName = App.Path
            If Right(fName, 1) <> "\" Then fName = fName + "\"
            fName = fName + "autosave.log"
            Open fName For Append As fileNum
                Print #fileNum, txt
            Close fName
            If (chkSaveAs.Value = 1) And (chkCurrent.Value = 0) Then
                myHWnd = CLng(List2.List(List1.ListIndex))
                lRetVal = SetForegroundWindow(myHWnd)
                ' SetForegroundWindow takes the equivilant hWnd from an invisible list
                ' and activates that window. However, if the window is minimized, then
                ' it is only activated in the taskbar, and not maximized
                bRetVal = OpenIcon(myHWnd)
                ' Maximizes the window
                SendKeys ("%FA")
                   
            ElseIf chkCurrent.Value = 1 Then
                SendKeys ("%FS")    ' This time we just save the file, and don't
                                    ' confirm file name.
                         
            ElseIf chkCurrent.Value = 1 Then
                SendKeys ("%FA")    ' This sends alt-F, then A to the current
                                    ' window (whichever window is in focus).
            
            Else
                myHWnd = CLng(List2.List(List1.ListIndex))
                lRetVal = SetForegroundWindow(myHWnd)
                bRetVal = OpenIcon(myHWnd)
                SendKeys ("%FS")
            End If
        txtRemaining = txtInterval.Text
        ProgressBar1.Value = 0
        End If
    End If
End Sub
Private Sub tmrRefresh_Timer()
    ' Checks if a list item is selected and if the timer is set
    If txtInterval = "" Then Exit Sub
    On Error Resume Next
    If (TimerSet = True Or txtInterval > 0 And txtInterval < "A") And List1.SelCount <> 0 Then
        optOn.Enabled = True
    Else
        optOn.Enabled = False
    End If
    
    ' Can only view or kill a file when autosave is off because otherwise the file
    ' might be open.
    If optOn.Value = True Then
        cmdKill.Enabled = False
    Else
        cmdKill.Enabled = True
    End If
    
End Sub

Private Sub tmrWait_Timer()
Dim i As Integer
Dim txt As String
Dim txt2 As Long

'Dim hwCurr As Long
Dim intLen As Long
Dim strTitle As String
       
'Enumerates open (including hidden) windows into a list box
    tmrWait.Enabled = False
      For i = 1 To NumProcs
        With ProcInfo(i)
            txt = .Title    ' List1 will get the Window Title
            txt2 = .AppHwnd ' List2 will get the Window's hWnd
        End With
        
        intLen = GetWindowTextLength(ProcInfo(i).AppHwnd) + 1
        strTitle = Space$(intLen)
        intLen = GetWindowText(hwCurr, strTitle, intLen)
        
        'Dump System and Blank Procs
        If chkVisible.Value = 1 Then
            If TaskWindow(ProcInfo(i).AppHwnd) Then 'If the window is visible
                List1.AddItem txt, 0
                List2.AddItem txt2, 0
            Else
                NumProcs = NumProcs - 1 ' Otherwise, the number of Procs is decreased
                                        ' by one
            End If
        Else 'Any open window, but we'll try to dump some of the obvious system ones:
            If (Left(txt, 1) <> " ") And (Left(txt, 2) <> "MS") And _
            (Left(txt, 3) <> "WIN") And (Left(txt, 2) <> "DD") And _
            (Left(txt, 3) <> "OLE") And (Left(txt, 3) <> "Ole") Then
                List1.AddItem txt, 0
                List2.AddItem txt2, 0
            Else
                NumProcs = NumProcs - 1 ' Otherwise, the number of Procs is decreased
                                        ' by one
            End If
        End If
    Next i
    If chkVisible = 1 Then List1.RemoveItem NumProcs - 1
End Sub
Private Sub tmrKey_Timer()
    diDEV.GetDeviceStateKeyboard diState    'get all the key states.
    DoEvents    'doevents. Lets windows do anything it needs to do. Required
    'otherwise you can get it doing more things than it's capable of.

'60=F2
'61=F3

If diState.Key(60) <> 0 Then
    If optOn.Enabled = True Then optOn.Value = True
End If
If diState.Key(61) <> 0 Then
    If optOn.Enabled = True Then optOff.Value = True
End If
End Sub
Private Sub txtInterval_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 5
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub

Private Sub txtRemaining_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.HelpContextID = 6
    If Button = 2 Then PopupMenu mnuPopup
    DoEvents
    Me.HelpContextID = 100

End Sub
