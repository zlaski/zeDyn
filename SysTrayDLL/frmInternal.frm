VERSION 5.00
Begin VB.Form frmInternal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About SystrayDll"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmInternal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   2640
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   1620
      TabIndex        =   0
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   4275
   End
   Begin VB.Menu mPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmInternal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdOk_Click()
    Hide
End Sub

Private Sub Form_Load()
    Label1.Caption = "This code was written by Alan Toews, June 14, 2001" & vbCrLf & _
    "Feel free to use or modify this code, but please" & vbCrLf & _
    "do not take credit for it. If you use , find a bug," & vbCrLf & _
    "or have a suggestion, please let me know." & vbCrLf & _
    "Feedback encourages development, and is one of the few" & vbCrLf & _
    "returns an author gets for distributing free code." & vbCrLf & _
    "Thanks for looking!" & vbCrLf & _
    "Alan." & vbCrLf & _
    "actoews@hotmail.com"
End Sub





Private Sub mAbout_Click()
    Show
End Sub


