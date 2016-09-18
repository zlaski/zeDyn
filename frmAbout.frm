VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3450
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2381.251
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet InetUpdate 
      Left            =   0
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdate 
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1050
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   737.45
      ScaleMode       =   0  'User
      ScaleWidth      =   632.1
      TabIndex        =   1
      Top             =   240
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2880
      Width           =   1260
   End
   Begin VB.Label lblze 
      Caption         =   "http://www.zoneedit.com/dynamic-dns"
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
      Left            =   1440
      MouseIcon       =   "frmAbout.frx":1A8A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1853.234
      Y2              =   1853.234
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   1410
      TabIndex        =   2
      Top             =   1530
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1410
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1863.588
      Y2              =   1863.588
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1410
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   240
      TabIndex        =   3
      Top             =   2745
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    On Error GoTo hErr
    Unload Me
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub


Private Sub Form_Load()
    On Error GoTo hErr
    Me.Caption = "About " & frmMain.Caption
    lblze.ToolTipText = lblze.Caption
    lblTitle.Caption = frmMain.Caption
    If RL = False Then
        lblVersion.Alignment = vbLeftJustify
        lblVersion.RightToLeft = False
        lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        lblTitle.Alignment = vbLeftJustify
        lblTitle.RightToLeft = False
        lblDescription.Alignment = vbLeftJustify
        lblDescription.RightToLeft = False
        lblDescription = "Includes SysTrayDLL control created by Alan Toews"
    Else
        lblVersion.Alignment = vbRightJustify
        lblVersion.RightToLeft = True
        lblVersion.Caption = "«·‰”Œ… : " & App.Major & "." & App.Minor & "." & App.Revision
        lblTitle.Alignment = vbRightJustify
        lblTitle.RightToLeft = True
        lblDescription.Alignment = vbRightJustify
        lblDescription.RightToLeft = True
        lblDescription = "Includes SysTrayDLL control created by Alan Toews"
    End If
    lblDisclaimer = "(C) 1999-2016 Zone Edit, LLC All Rights Reserved."
    Me.Top = (Screen.Height * 0.85) / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Private Sub lblze_Click()
    On Error GoTo hErr
    ShellDocument lblze.Caption
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub
