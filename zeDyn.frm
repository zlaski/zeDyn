VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2670
   ClientLeft      =   7080
   ClientTop       =   6600
   ClientWidth     =   6375
   Icon            =   "zeDyn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   2640
      Top             =   2400
   End
   Begin VB.FileListBox File 
      Height          =   2235
      Left            =   10680
      Pattern         =   "*.lng"
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox Lst 
      Height          =   1230
      ItemData        =   "zeDyn.frx":030A
      Left            =   6480
      List            =   "zeDyn.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   3240
      Width           =   3735
   End
   Begin zeDyn.utimerCTL uTick 
      Left            =   0
      Top             =   120
      _ExtentX        =   635
      _ExtentY        =   582
   End
   Begin VB.HScrollBar hsTime 
      Height          =   255
      Left            =   7800
      Max             =   120
      Min             =   1
      TabIndex        =   15
      Top             =   2520
      Value           =   18
      Width           =   2415
   End
   Begin VB.CommandButton cmdAbout 
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdsh 
      Height          =   2415
      Left            =   6000
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox txtLog 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   3240
      Width           =   5655
   End
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet uInet 
      Left            =   2760
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox uHosts 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   960
      Width           =   5655
   End
   Begin VB.TextBox uPass 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.TextBox uLogin 
      Height          =   285
      Left            =   6600
      TabIndex        =   4
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton uOK 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblUrl 
      Height          =   495
      Left            =   6480
      TabIndex        =   24
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label lblMin 
      Height          =   255
      Left            =   6960
      TabIndex        =   23
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblmytime 
      Alignment       =   2  'Center
      Caption         =   "The Time is"
      Height          =   195
      Left            =   2400
      TabIndex        =   22
      Top             =   0
      Width           =   1155
   End
   Begin VB.Label lblFP 
      Caption         =   "Forget Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6480
      MouseIcon       =   "zeDyn.frx":030E
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblCreate 
      Caption         =   "Create New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6480
      MouseIcon       =   "zeDyn.frx":0FD8
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label lblLang 
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label lblTime 
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblInterval 
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblLog 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   5655
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5655
   End
   Begin VB.Label uStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label lblPass 
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label lblUser 
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblHost 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5655
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "&Check Now"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' tray code
Private inTray As Boolean
Dim WithEvents cTray As SysTrayDll.SysTray
Attribute cTray.VB_VarHelpID = -1

'### Declarations - place this part in the Declarations section of a code file (at the very top of the code)
Public Enum eSpecialFolders
  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
End Enum
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type


'### The routine - place this part before/after any other rotuine
Public Function SpecialFolder(CSIDL As Long) As String
    Dim sPath As String
    Dim IDL As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the "Recent Documents" folder.
    ' Info is stored in the IDL structure.
    '
    SpecialFolder = ""
    If SHGetSpecialFolderLocation(frmMain.hwnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            SpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
        End If
    End If
End Function


Private Sub cmdAbout_Click()
    On Error GoTo hErr
    frmAbout.Show vbModal
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub cmdExit_Click()
    On Error GoTo hErr
    Dim Msgr As String
    If RL = True Then
        Msgr = MsgBox(ReadLng("Messages", "exit", LngFile), vbCritical + vbOKCancel + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title)
    Else
        Msgr = MsgBox(ReadLng("Messages", "exit", LngFile), vbCritical + vbOKCancel, App.Title)
    End If
    If Msgr = vbOK Then
        If Me.WindowState = vbMinimized Then
            Me.Visible = True
            Me.WindowState = vbNormal
            Me.Show
        End If
        WriteLog ReadLng("Messages", "stop", LngFile)
        Set cTray = Nothing
        inTray = False
        End
    End If
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub cmdsh_Click()
    On Error GoTo hErr
    If Width = 6465 Then 'Less Info
        Me.Visible = True
        cTray.FormRestore
        Width = 10425
        Height = 5745
        cmdAbout.Top = 4800
        cmdExit.Top = 4800
        uOK.Top = 4800
        lblLog.Top = 2040
        txtLog.Top = 2400
        cmdsh.Caption = ReadLng("Main", "LI", LngFile)
        cmdsh.Top = 1440
        Me.Top = (Screen.Height * 0.85) / 2 - Me.Height / 2
        Me.Left = Screen.Width / 2 - Me.Width / 2
    ElseIf Width = 10425 Then 'More Info
        Me.Visible = True
        cTray.FormRestore
        Width = 6465
        Height = 3105
        cmdAbout.Top = 2160
        cmdExit.Top = 2160
        uOK.Top = 2160
        lblLog.Top = 3940
        txtLog.Top = 3800
        cmdsh.Caption = ReadLng("Main", "MI", LngFile)
        cmdsh.Top = 120
        Me.Top = (Screen.Height * 0.85) / 2 - Me.Height / 2
        Me.Left = Screen.Width / 2 - Me.Width / 2
    End If

Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub cTray_DblClick(button As Integer)
    On Error GoTo hErr
    Call mnuOpen_Click
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
    On Error GoTo hErr
    Set cTray = New SysTrayDll.SysTray
    Dim CommonAppData As String
    
    CommonAppData = SpecialFolder(SpecialFolder_CommonAppData) & "\ZoneEdit Dynamic DNS Update Client\"
    MakeSureDirectoryPathExists CommonAppData
    
    gINIFile = CommonAppData + "zeDyn.ini"
    LoadINI
    
    gLogFile = CommonAppData + "zeDyn.log"
    Caption = ReadLng("Main", "Caption", LngFile) & " " & App.Major & "." & App.Minor
    
    lblTime.Caption = hsTime.Value
    uTick.MinuteInterval = lblTime.Caption
    
    cTray.Form = Me
    cTray.PopupMenu = mnuTray
    cTray.PopupStyle = stOnRightUp
    cTray.RestoreFromTrayOn = stOnLeftDblClick + stOnRightDblClick
    cTray.Icon = Me.Icon
    cTray.TrayTip = Me.Caption
    cTray.Visible = True
    
    Dim i As Integer
    File.Path = App.Path
    For i = 0 To File.ListCount - 1
        Lst.AddItem (Mid(File.List(i), 1, Len(File.List(i)) - 4))
        ' Highlight the startup language file
        If InStr(1, LngFile, File.List(i)) Then
            Selitmlst = i
        End If
    Next
    If Lst.ListCount = 0 Then
        If RL = True Then
            MsgBox ReadLng("Messages", "nolangfile", LngFile), vbCritical + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title
        Else
            MsgBox ReadLng("Messages", "nolangfile", LngFile), vbCritical, App.Title
        End If
        Unload Me
    Else
        Lst.Selected(Selitmlst) = True
    End If
    
    uPass.PasswordChar = Chr(&H95)
    
    Me.Top = (Screen.Height * 0.85) / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2

    If uLogin.Text <> "" And uPass.Text <> "" And uHosts.Text <> "" Then
        Call uOK_Click
    Else
        Call cmdsh_Click
        uStatus.Caption = ReadLng("Messages", "checkinf", LngFile)
        frmMain.uStatus.ForeColor = COLOR_RED
    End If
    
    WriteLog ReadLng("Messages", "start", LngFile)
    DoEvents
    ReadLog
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo hErr
    Call uOK_Click
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub hsTime_Change()
    On Error GoTo hErr
    lblTime = hsTime.Value
    uTick.Interval = lblTime
    WriteINIString gINIFile, "Settings", "Time", lblTime
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub lblCreate_Click()
    On Error GoTo hErr
    ShellDocument "http://www.zoneedit.com/signup.html"
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub lblFP_Click()
    On Error GoTo hErr
    ShellDocument "http://www.zoneedit.com/lostpass.html"
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub Lst_Click()
    On Error GoTo hErr
    Selitmlst = Lst.ListIndex
    LngFile = App.Path & "\" & Lst.Text & ".lng"
    ReadLanguage
    WriteINIString gINIFile, "Settings", "Lang", LngFile
    CheckData
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub mnuAbout_Click()
    On Error GoTo hErr
    frmAbout.Show vbModal
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub mnuCheck_Click()
    On Error GoTo hErr
    uTick_Timer
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub mnuOpen_Click()
    On Error GoTo hErr
    Me.Visible = True
    cTray.FormRestore
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub mnuQuit_Click()
    On Error GoTo hErr
    Call cmdExit_Click
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub tmrTime_Timer()
    On Error GoTo hErr
    lblmytime.Caption = Time

Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub uOK_Click()
    On Error GoTo hErr
    WriteINIString gINIFile, "Settings", "Login", uLogin.Text
    WriteINIString gINIFile, "Settings", "Pass", uPass.Text
    
    Dim HostList As String, UnpackedHosts As String
    HostList = PackHosts(uHosts.Text)
    UnpackedHosts = UnpackHosts(HostList)
    
    If uHosts.Text <> UnpackedHosts Then
        uHosts.Text = UnpackedHosts
        gCurIP = ""
        WriteINIString gINIFile, "Settings", "IP", ""
    End If
    
    WriteINIString gINIFile, "Settings", "Hosts", HostList
    
    Me.Visible = False
    cTray.FormMinimize
    
    uTick_Timer
    CheckData
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub uStatus_Click()
Select Case uStatus.Tag
    Case "ipuf"
        uOK_Click
    Case "ipdf"
        uOK_Click
End Select
End Sub

Private Sub uTick_Timer()
    On Error GoTo ErrTime
    Dim ip As String
    Dim i As Integer
    
    If CheckData = False Then Exit Sub
    
    'If Me.WindowState = vbNormal Then Exit Sub
    
    If uInet.StillExecuting Then
        DoEvents
        uOK_Click
        Exit Sub
    End If
    
    ' get IP
    ip = uInet.OpenURL(gIPDetectURL, icString)
    i = InStr(ip, gIPDetectPrefix)
    
    If i >= 1 Then
        ' This is probably not needed anymore.  ZoneEdit now just returns the bare IP address.
        ip = Mid(ip, Len(gIPDetectPrefix))
        i = InStr(ip, "<")
        
        If i < 1 Then
            IPError
            Exit Sub
        End If
        
        ip = Left(ip, i - 1)
    End If
    ip = Trim(ip)
    
    If ip <> gCurIP Then
        ' Update DNS
        Dim ok As String
        Dim hosts As String
        hosts = PackHosts(uHosts.Text)
        
        uStatus.Caption = "Changing IP to " & ip
        uStatus.ForeColor = COLOR_BLUE
        
        uInet.URL = "https://dynamic.zoneedit.com/auth/dynamic.html?dnsto=" & ip & "&host=" & hosts
        uInet.UserName = uLogin.Text
        uInet.Password = uPass.Text
        ok = uInet.OpenURL
        
        ok = Replace(ok, vbCr, " ")
        ok = Replace(ok, vbLf, " ")

        WriteLog ok

        If InStr(ok, "<SUCCESS") Then
            gCurIP = ip
            uStatus.Caption = gCurIP
            uStatus.ForeColor = COLOR_GREEN
            
            WriteINIString gINIFile, "Settings", "IP", ip
            
            Dim tip As String
            cTray.TrayTip = "ZoneEdit Dynamic DNS Update Client" & " (" & gCurIP & ")"
            
        Else
            uStatus.Caption = ReadLng("Messages", "ipuf", LngFile)
            uStatus.ForeColor = COLOR_RED
            uStatus.Tag = "ipuf"
        End If
    End If
    ReadLog
    
ErrTime:
    If Err.Number = 35761 Then
        WriteLog ReadLng("Messages", "firwall", LngFile)
        uStatus.Caption = ReadLng("Messages", "firewall", LngFile)
        uStatus.ForeColor = COLOR_RED
        ReadLog
    ElseIf Err.Number = 0 Then
        DoEvents
    Else
        MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    End If
End Sub

Private Sub IPError()
    On Error GoTo hErr
    WriteLog ReadLng("Messages", "ipdf", LngFile)
    uStatus.Caption = ReadLng("Messages", "ipdf", LngFile)
    uStatus.ForeColor = COLOR_RED
    ReadLog
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

' Get rid of all manner of spaces and replace them with single commas
Function PackHosts(ByVal hosts As String)
    On Error GoTo hErr
    ' Allow arbitrary spaces/commas between host/domain names
    hosts = Replace(hosts, " ", ",")
    hosts = Replace(hosts, vbCrLf, ",")
    hosts = Replace(hosts, vbCr, ",")
    hosts = Replace(hosts, vbLf, ",")
    Do While InStr(hosts, ",,")
        hosts = Replace(hosts, ",,", ",")
    Loop
    If Left(hosts, 1) = "," Then
        hosts = Mid(hosts, 2)
    End If
    If Right(hosts, 1) = "," Then
        hosts = Mid(hosts, 1, Len(hosts) - 1)
    End If
    
    PackHosts = hosts
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

' Make each domain/host appear on a separate line in the TextBox
Function UnpackHosts(ByVal hosts As String)
    On Error GoTo hErr
    ' Allow arbitrary spaces/commas between host/domain names
    hosts = PackHosts(hosts)
    hosts = Replace(hosts, ",", vbCrLf)
    UnpackHosts = hosts
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

Private Sub LoadINI()
    On Error GoTo hErr
    uLogin.Text = GetINIString(gINIFile, "Settings", "Login", "")
    uPass.Text = GetINIString(gINIFile, "Settings", "Pass", "")
    uHosts.Text = UnpackHosts(GetINIString(gINIFile, "Settings", "Hosts", ""))
    gIPDetectURL = GetINIString(gINIFile, "Settings", "IPDetectURL", "https://dynamic.zoneedit.com/checkip.html")
    gIPDetectPrefix = GetINIString(gINIFile, "Settings", "IPDetectPrefix", "IP Address:")
    LngFile = GetINIString(gINIFile, "Settings", "Lang", App.Path & "\English.lng")
    hsTime.Value = Int(GetINIString(gINIFile, "Settings", "Time", "30"))
    ReadLanguage
    
    gCurIP = GetINIString(gINIFile, "Settings", "IP", "")
    uStatus.Caption = gCurIP
    uStatus.ForeColor = COLOR_GREEN
    
    ReadLog
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub
