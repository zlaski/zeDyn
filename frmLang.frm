VERSION 5.00
Begin VB.Form frmLang 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6600
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   11115
   ControlBox      =   0   'False
   Icon            =   "frmLang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11115
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File 
      Height          =   2235
      Left            =   6360
      Pattern         =   "*.lng"
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox Lst 
      BackColor       =   &H00404080&
      ForeColor       =   &H0000FFFF&
      Height          =   2400
      ItemData        =   "frmLang.frx":030A
      Left            =   2640
      List            =   "frmLang.frx":030C
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmLang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub cmdOk_Click()
    frmMain.Enabled = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Selitmlst = Int(GetINIString(gINIFile, "Startup options", "SelectOne", "0"))
    Lst.Selected(Selitmlst) = True
    LngFile = App.Path & "\" & Lst.Text & ".lng"
    frmMain.Enabled = True
    Unload Me
End Sub

Private Sub Form_Click()
SaveLanguage
End Sub

Private Sub Form_Load()
    'Load last selected language
    File.Path = App.Path
    For i = 0 To File.ListCount - 1
        Lst.AddItem (Mid(File.List(i), 1, Len(File.List(i)) - 4))
    Next
    If Lst.ListCount = 0 Then
        MsgBox ReadLng("Messages", "nolangfile", LngFile), vbCritical + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title
        Unload Me
    Else
        Lst.Selected(Selitmlst) = True
    End If
    Selitmlst = Int(GetINIString(gINIFile, "Startup options", "SelectOne", "0"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
    SaveSetting App.Title, "Startup options", "SelectOne", Selitmlst
    Unload Me
End Sub
