VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmTest 
   Caption         =   "Test Form"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   1680
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":159A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":19EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Persistent Icon"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Icon will be restored if explorer crashes"
      Top             =   480
      Width           =   1575
   End
   Begin VB.HScrollBar hsIcons 
      Height          =   255
      Left            =   2760
      Min             =   1
      TabIndex        =   6
      Top             =   720
      Value           =   1
      Width           =   735
   End
   Begin VB.PictureBox pIcon 
      Height          =   735
      Left            =   2760
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Minimize to tray"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show In tray?"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmTest.frx":1E3E
      ToolTipText     =   "Type your tip, then hover the mouse over the tray icon."
      Top             =   1320
      Width           =   3375
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "frmTest.frx":1E52
      Left            =   120
      List            =   "frmTest.frx":1E54
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Tray Icon"
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Tray Tip: Type your tip, then hover the mouse over the tray icon."
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2475
   End
   Begin VB.Menu mPopup 
      Caption         =   "Popup"
      Begin VB.Menu mRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "About SysTrayDll"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents cTray As SysTrayDll.SysTray
Attribute cTray.VB_VarHelpID = -1


Private Sub Check1_Click()
    cTray.Form = Me
    cTray.PopupMenu = mPopup
    cTray.PopupStyle = stOnRightUp
    
    'when the icon is double clicked (left or right mouse), the form will restore.
    cTray.RestoreFromTrayOn = stOnLeftDblClick + stOnRightDblClick
    
    'cTray.Icon = Me.Icon
    cTray.TrayTip = Text1.Text
    cTray.Visible = CBool(Check1.Value)
End Sub





Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        cTray.TrayFormStyle = stHideFormWhenMin + stHideTrayWhenNotMin
    Else
        cTray.TrayFormStyle = stNormal
        cTray.Visible = CBool(Check1.Value)
    End If
End Sub

Private Sub Check3_Click()
    cTray.Persistent = CBool(Check3)
End Sub

Private Sub cTray_Click(button As Integer)
    List1.AddItem ButtonName(button) & " Click"
End Sub

Private Sub cTray_DblClick(button As Integer)
    List1.AddItem ButtonName(button) & " Double-Click"
End Sub
Private Function ButtonName(button As Integer) As String
    Select Case button
        Case vbLeftButton
            ButtonName = "Left"
        Case vbRightButton
            ButtonName = "Right"
        Case vbMiddleButton
            ButtonName = "Middle"
    End Select
End Function

Private Sub cTray_MouseDown(button As Integer)
    List1.AddItem ButtonName(button) & " Mouse Down"
End Sub



Private Sub cTray_Refreshed()
    Check1 = Abs(cTray.Visible)
    MsgBox "Explorer has been restarted."
End Sub

Private Sub Form_Load()
    Set cTray = New SysTrayDll.SysTray
    Text1.Text = "Type your tray tip here." & vbCrLf & _
                 "It can be up to 64 characters long" & vbCrLf & _
                 "counting hidden chars"
                 
    pIcon.Picture = cTray.Icon
    Check3 = Abs(cTray.Persistent)
    
    ilIcons.ListImages.Add , , cTray.Icon
    hsIcons.Max = ilIcons.ListImages.Count
    
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set cTray = Nothing
End Sub

Private Sub hsIcons_Change()
    pIcon.Picture = ilIcons.ListImages(hsIcons.Value).Picture
    cTray.Icon = ilIcons.ListImages(hsIcons.Value).Picture
End Sub

Private Sub List1_Click()
    
    List1.Selected(List1.ListCount - 1) = True
    
End Sub

Private Sub mAbout_Click()
    cTray.ShowAbout
End Sub

Private Sub mExit_Click()
    Unload frmTest
End Sub

Private Sub mRestore_Click()
    cTray.FormRestore
End Sub

Private Sub Text1_Change()
    cTray.TrayTip = Text1.Text
    
End Sub


