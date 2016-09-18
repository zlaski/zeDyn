VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This code was written by Alan Toews, June 14, 2001
'Feel free to use or modify this code, but please
'do not take credit for it. If you use , find a bug,
'or have a suggestion, please let me know.
'Feedback encourages development, and is one of the few
'returns an author gets for distributing free code.
'You can contact me at:
'
'actoews@hotmail.com


Option Explicit

'apis for detecting explorer crash
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
'this type is necessary for the Shell_NotifyIcon API
Private Type NOTIFYICONDATA
 cbSize As Long
 hwnd As Long
 uId As Long
 uFlags As Long
 uCallBackMessage As Long
 hIcon As Long
 szTip As String * 64
End Type

'constants for use with SetWindowLong API
Private Const GWL_WNDPROC = (-4)
Private Const GWL_USERDATA = (-21)

'Constants for use with Shell_NotifyIcon API
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'Constants for use with NOTIFYICONDATA Type
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'Constants for use in Callback lParam value
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBCLICK = &H209

'API declarations for setting the tray icon
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'API declarations for capturing mouse events in the tray icon
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Variables used to set & clear
'tray icon & callback functions
Private nidTray As NOTIFYICONDATA
Private WindowProc As Long

'public events to be raised when the user interacts with the icon
Public Event MouseDown(button As Integer)
Public Event MouseUp(button As Integer)
Public Event MouseMove()
Public Event Click(button As Integer)
Public Event DblClick(button As Integer)
Public Event Minimize()
Public Event Restore()
Public Event Refreshed()

'enumeration type used to set the tray click behaviors.
Public Enum stEnumTrayMenu
    stNone = &H0
    stOnLeftUp = &H1
    stOnRightUp = &H2
    stOnMiddleUp = &H4
    stOnLeftDown = &H8
    stOnRightDown = &H10
    stOnMiddleDown = &H20
    stOnLeftDblClick = &H40
    stOnRightDblClick = &H80
    stOnMiddleDblClick = &H100
    stOnAllClickEvents = &H1FF
End Enum

'enumeration to set tray icon behavior
Public Enum stEnumTrayStyle
    stHideFormWhenMin = 1
    stHideTrayWhenNotMin = 2
    stNormal = 0
End Enum

'Private property variables
Private stIcon As StdPicture    'Icon to use for the tray
Private stVisible As Boolean    'Is icon shown in the tray or not
Private WithEvents stForm As Form          'Parent form (needed for hWnd)
Attribute stForm.VB_VarHelpID = -1
Private WithEvents stTimer As Timer        'Timer object used for persistence
Attribute stTimer.VB_VarHelpID = -1
Private stMenu As Menu          'right click menu
Private sthWnd As Long
Private stTrayTip As String     'tool tip to appear when mouse is hovering over tray
Private stMenuStyle As stEnumTrayMenu
Private stTrayStyle As stEnumTrayStyle
Private stRestoreFromTray As stEnumTrayMenu
Private stPersistent As Boolean

'private variables
Private LastWindowState
Private LastTrayHWND As Long
'*********************************************
'Declare all properties here:
'*********************************************
Public Property Get Persistent() As Boolean
    Persistent = stTimer.Enabled
End Property
Public Property Let Persistent(NewVal As Boolean)
    stTimer.Interval = 1000
    stTimer.Enabled = NewVal
End Property
Public Property Get RestoreFromTrayOn() As stEnumTrayMenu
    RestoreFromTrayOn = stRestoreFromTray
End Property
Public Property Let RestoreFromTrayOn(NewVal As stEnumTrayMenu)
    stRestoreFromTray = NewVal
End Property
Public Property Get TrayFormStyle() As stEnumTrayStyle
    TrayFormStyle = stTrayStyle
End Property
Public Property Let TrayFormStyle(NewVal As stEnumTrayStyle)
    stTrayStyle = NewVal
    If CBool(stTrayStyle And stHideTrayWhenNotMin) Then If stForm.Visible Then RemoveFromTray
End Property

Public Property Get PopupStyle() As stEnumTrayMenu
    PopupStyle = stMenuStyle
End Property
Public Property Let PopupStyle(NewVal As stEnumTrayMenu)
    stMenuStyle = NewVal
End Property
Public Property Get Icon() As StdPicture
    Set Icon = stIcon
End Property
Public Property Let Icon(NewVal As StdPicture)
    'change the icon in the tray
    Set stIcon = NewVal
    
    'if the tray icon is loaded, then change it now.
    If Not stVisible Then Exit Property
    ' Update the tray icon.
    With nidTray
        .hIcon = stIcon.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, nidTray
End Property
Public Property Get Visible() As Boolean
    Visible = stVisible
End Property
Public Property Let Visible(NewVal As Boolean)
    'toggle whether the icon is loaded in the form or not
    ' the stvisible variable is set in each of the called functions
    'also, if we've set to hide the tray icon when the tray is visible, then
    'check that the form is not visible before showing
    
    If CBool(stTrayStyle And stHideTrayWhenNotMin) Then NewVal = Not stForm.Visible
    If NewVal = stVisible Then Exit Property
    If NewVal Then
        ShowInTray
    Else
        RemoveFromTray
    End If
End Property
Public Property Get Form() As Object
    Set Form = stForm
End Property
Public Property Let Form(NewVal As Object)
    'if the form changes, then make sure to
    'clear any popup menu specified.
    'if the user tries to call a popup menu on a form other than
    'the stForm form, then we'll get an error, or
    'the wrong form will be displayed.
    If NewVal Is Nothing Then Set NewVal = frmInternal
    
    If NewVal.Name <> stForm.Name Then
        Set stForm = NewVal
        Set stMenu = Nothing
        LastWindowState = stForm.WindowState
        If LastWindowState = vbMinimized Then LastWindowState = vbNormal
    End If
End Property
Public Property Get PopupMenu() As Object
    Set PopupMenu = stMenu
End Property
Public Property Let PopupMenu(NewVal As Object)
    Set stMenu = NewVal
End Property
Public Property Get TrayTip() As String
    TrayTip = stTrayTip
End Property
Public Property Let TrayTip(NewVal As String)
    stTrayTip = NewVal
    
    'if the icon is in the tray, then
    'update the tip immediately
    If Not stVisible Then Exit Property
    With nidTray
        .szTip = stTrayTip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, nidTray
End Property

'*********************************************
'Declare all private subs & functions here:
'*********************************************
Private Sub ShowInTray()
    
    'On Error Resume Next
    With nidTray
        .cbSize = Len(nidTray)
        .hwnd = stForm.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = stIcon.Handle
        .szTip = stTrayTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nidTray
    stVisible = True
    
    'set tray callback function
    WindowProc = SetWindowLong(stForm.hwnd, GWL_WNDPROC, AddressOf mCallbackFunction)
End Sub

Private Sub RemoveFromTray()
    'Disable callback
    SetWindowLong stForm.hwnd, GWL_WNDPROC, WindowProc
    
    'If Not stVisible Then Exit Sub
    Shell_NotifyIcon NIM_DELETE, nidTray
    stVisible = False

End Sub

Private Sub Class_Initialize()
    'grab the hWnd of the tray to make sure it is loaded
    LastTrayHWND = FindWindow("Shell_TrayWnd", vbNullString)
    
    'make sure that the form and menu variables point to somewhere.
    Set stForm = frmInternal
    Set stTimer = frmInternal.Timer1
    Set stIcon = frmInternal.Icon
    
    stMenuStyle = stOnRightUp
    Set stMenu = frmInternal.mPopup
    Init Me
    stRestoreFromTray = stOnLeftDblClick
End Sub

Private Sub Class_Terminate()
    'clean up behind us
    On Error Resume Next
    RemoveFromTray
    Set stForm = Nothing
    Set stIcon = Nothing
    Unload frmInternal
End Sub

Friend Function CallBack(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
        'if the user clicked on the tray icon then
        'lets look at how they clicked and decide what to do.
        'when we set the tray icon, we gave it a callback message value of WM_MOUSEMOVE
        'if we receive that value, then we know its the tray icon responding.
        
        'the lParam value specifies what action was performed.
        'ie: left mouse down, middle mouse double click, etc..
        If Msg = WM_MOUSEMOVE Then
            'if no menu is specified, then set the display style to none
            If stMenu Is Nothing Then stMenuStyle = stNone
            'for each case we want to check whether the menu should be displayed or not
            Select Case lParam
                'case button is down
                Case WM_LBUTTONDOWN
                    If CBool(stRestoreFromTray And stOnLeftDown) And stForm.WindowState = vbMinimized Then FormRestore
                    RaiseEvent MouseDown(vbLeftButton)
                    If CBool(stMenuStyle And stOnLeftDown) Then stForm.PopupMenu stMenu
                Case WM_RBUTTONDOWN
                    If CBool(stRestoreFromTray And stOnRightDown) And stForm.WindowState = vbMinimized Then FormRestore
                    RaiseEvent MouseDown(vbRightButton)
                    If CBool(stMenuStyle And stOnRightDown) Then stForm.PopupMenu stMenu
                Case WM_MBUTTONDOWN
                    If CBool(stRestoreFromTray And stOnMiddleDown) And stForm.WindowState = vbMinimized Then FormRestore
                    RaiseEvent MouseDown(vbMiddleButton)
                    If CBool(stMenuStyle And stOnMiddleDown) Then stForm.PopupMenu stMenu
                    
                'case button double click
                Case WM_LBUTTONDBLCLK
                    If CBool(stRestoreFromTray And stOnLeftDblClick) And stForm.WindowState = vbMinimized Then FormRestore
                    If CBool(stMenuStyle And stOnLeftDblClick) Then stForm.PopupMenu stMenu
                    RaiseEvent DblClick(vbLeftButton)
                Case WM_RBUTTONDBLCLK
                    If CBool(stRestoreFromTray And stOnRightDblClick) And stForm.WindowState = vbMinimized Then FormRestore
                    RaiseEvent DblClick(vbRightButton)
                    If CBool(stMenuStyle And stOnRightDblClick) Then stForm.PopupMenu stMenu
                Case WM_MBUTTONDBCLICK
                    If CBool(stRestoreFromTray And stOnMiddleDblClick) And stForm.WindowState = vbMinimized Then FormRestore
                    RaiseEvent DblClick(vbMiddleButton)
                    If CBool(stMenuStyle And stOnMiddleDblClick) Then stForm.PopupMenu stMenu
                
                'case button up
                Case WM_LBUTTONUP
                    If CBool(stRestoreFromTray And stOnLeftUp) And stForm.WindowState = vbMinimized Then FormRestore
                    RaiseEvent MouseUp(vbLeftButton)
                    RaiseEvent Click(vbLeftButton)
                    If CBool(stMenuStyle And stOnLeftUp) Then stForm.PopupMenu stMenu
                Case WM_RBUTTONUP
                    If CBool(stRestoreFromTray And stOnRightUp) And stForm.WindowState = vbMinimized Then FormRestore
                    RaiseEvent MouseUp(vbRightButton)
                    RaiseEvent Click(vbRightButton)
                    If CBool(stMenuStyle And stOnRightUp) Then stForm.PopupMenu stMenu
                Case WM_MBUTTONUP
                    If CBool(stRestoreFromTray And stOnMiddleUp) And stForm.WindowState = vbMinimized Then FormRestore
                    RaiseEvent MouseUp(vbMiddleButton)
                    RaiseEvent Click(vbMiddleButton)
                    If CBool(stMenuStyle And stOnMiddleUp) Then stForm.PopupMenu stMenu
                    
                'case mouse moves
                Case WM_MOUSEMOVE
                    RaiseEvent MouseMove
            End Select
        End If
        
        CallBack = CallWindowProc(WindowProc, hwnd, Msg, wParam, lParam)
End Function

Private Sub stForm_Resize()
    If stForm.WindowState <> vbMinimized Then
        If Not stForm.Visible Then stForm.Visible = True
        LastWindowState = stForm.WindowState
    Else
        FormMinimize
    End If
End Sub

'*********************************************
'Declare all public subs & functions here:
'*********************************************
Public Sub ShowAbout()
    frmInternal.Show
End Sub
Public Sub FormMinimize()
    
    'stForm.Visible = True
    If stForm.WindowState <> vbMinimized Then stForm.WindowState = vbMinimized
    
    If CBool(stTrayStyle And stHideFormWhenMin) Then stForm.Visible = False
    If CBool(stTrayStyle And stHideTrayWhenNotMin) And Not stVisible Then ShowInTray
    RaiseEvent Minimize
End Sub
Public Sub FormRestore()
    
    'for some reason, this needs to be run more than once
    'I put in the loop, so that it will ensure that the
    'form will be restored on the first try by the user.
    Do While stForm.WindowState = vbMinimized
        '** the form must be visible before we resize it, or we will crash.
        stForm.Visible = True
        'safety check in case the LastWindowstate somehow got set to minimized
        If LastWindowState = vbMinimized Then LastWindowState = vbNormal
        stForm.WindowState = LastWindowState
        
    Loop
    stForm.SetFocus
    'if desired, remove the icon from the tray when restored.
    If CBool(stTrayStyle And stHideTrayWhenNotMin) And stVisible Then RemoveFromTray
    RaiseEvent Restore
End Sub


Private Sub stTimer_Timer()
    'this function checks the hWnd of the system tray
    'if the value changes, then we will reload the tray icon
    Dim tmp As Long, x As Long
    
    'get the hWnd value of the system tray
    tmp = FindWindow("Shell_TrayWnd", vbNullString)
    
    
    'check for a change from the last time, and make sure it is a valid hWnd
    If (tmp <> LastTrayHWND) And (tmp > 0) Then
        Debug.Print stVisible, stPersistent
        If stVisible Then
            
            'reset subclassing callback
            RemoveFromTray
            
            'replace the icon in the system tray once explorer has restarted
            ShowInTray
            RaiseEvent Refreshed

        End If
    End If
    
    LastTrayHWND = tmp
End Sub
