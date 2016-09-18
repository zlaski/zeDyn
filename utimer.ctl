VERSION 5.00
Begin VB.UserControl utimerCTL 
   BackColor       =   &H00C0FFFF&
   CanGetFocus     =   0   'False
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2775
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   930
   ScaleWidth      =   2775
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "utimerCTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*UltraTimer ActiveX Control 1.1 - SOURCE CODE -***
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.
'Code Written by Greg Miller.
'PacZero (http://www.paczero.cjb.net)
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.
'This code is provided at no cost and can be
'modified and/or redistributed royalty free as long as
'the control's name has been changed.
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.
'**************************************************
'*****This usercontrol requires the VB Timer control
'*****which does the normal event handling.
'**************************************************
'*****Load this control module in your project by menu selecting
'*****Project/Add User Control... and select utimer.ctl
'**************************************************
'**************************************************
'Unlike VB's timer control which has a max interval
'of just over 1 minute, and requires it to be in milliseconds,
'the 'UltraTimer Control let's you combined this
'interval with additional 'minute and hour intervals,
'up to 24 days. At VB Design or 'Runmode, set the
'TimerMode to run as a One-time(single event)
'or as a periodic timer.
'**************************************************
'**************************************************
Option Explicit
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.
'Enum Property Values
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.
Enum eTimerMode
 Periodic
 OneTime
End Enum
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.
'Property Variables
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.
Dim m_mSeconds As Long
Dim m_Minutes As Integer
Dim m_Hours As Integer
Dim m_TimerMode As eTimerMode
Dim m_Enabled As Boolean
Dim MinutesElapsed As Integer
Dim msFlag As Boolean
Dim OneTimeUsed As Boolean
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'
'Event Declarations
'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'.'
Event Timer()
Event HourAlarm(HourMark As Integer)
Event MinuteAlarm(MinuteMark As Integer)
Public Property Let TimerMode(ByVal pVal As eTimerMode)
 On Error GoTo TimerModeErr
 
 'TimerMode has been changed
 Select Case pVal
  Case Periodic, OneTime 'Periodic or one-Time events
  Case Else
    MsgBox Error(380)
 End Select
 
 If m_TimerMode <> pVal Then
  m_TimerMode = pVal
  PropertyChanged "TimerMode"
 End If
 
 If OneTimeUsed And (Ambient.UserMode = True) Then ResetTimer
 
 Exit Property
 
TimerModeErr:     MsgBox Err.Description
End Property
Public Property Get TimerMode() As eTimerMode
  TimerMode = m_TimerMode
End Property
Public Property Get Interval() As Long
  Interval = m_mSeconds
End Property
Public Property Get MinuteInterval() As Integer
  MinuteInterval = m_Minutes
End Property
Public Property Get HourInterval() As Integer
  HourInterval = m_Hours
End Property
Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property
Public Property Let HourInterval(ByVal New_Hours As Integer)
Dim MaxHours As Long
If m_Hours <> New_Hours Then
 
 'Test for exceeded long value
 On Local Error GoTo MaxHours_err
 MaxHours& = ((New_Hours * 60) + m_Minutes) * 60000#
 
 'Hours has been changed
 m_Hours = New_Hours
 
 PropertyChanged "HourInterval"
 
 If (Ambient.UserMode = True) Then ResetTimer
End If
Exit Property
MaxHours_err:
MsgBox Error(380)
Exit Property
End Property
Public Property Let MinuteInterval(ByVal New_Minutes As Integer)
Dim MaxMins As Long
If m_Minutes <> New_Minutes Then
 
 ' Test for overflow and invalid property value
 On Local Error GoTo overflow_err
 MaxMins& = ((m_Hours * 60) + New_Minutes) * 60000#
 
 If New_Minutes >= 60 Then Exit Property
 
 'Minutes has been changed
 m_Minutes = New_Minutes
 PropertyChanged "MinuteInterval"
 
 If (Ambient.UserMode = True) Then ResetTimer
End If
Exit Property
overflow_err:
MsgBox Error(380)
Exit Property
End Property
Public Property Let Interval(ByVal New_mSeconds As Long)
If m_mSeconds <> New_mSeconds Then
 'mSeconds has been changed
 If New_mSeconds >= 60000 Then Exit Property
 m_mSeconds = New_mSeconds
 
 PropertyChanged "Interval"
 
 If (Ambient.UserMode = True) Then ResetTimer
End If
 
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
If New_Enabled <> m_Enabled Then
 'Enabled has been changed
 m_Enabled = New_Enabled
 
 PropertyChanged "Enabled"
 
 If (Ambient.UserMode = True) Then ResetTimer
End If
End Property
Private Sub Timer1_Timer()
 With Timer1
 
 'Stop current timer interval
 If Not .Enabled Then Exit Sub
 .Enabled = False
 
 If Not msFlag And m_mSeconds > 0 Then
 msFlag = True
 .Interval = m_mSeconds
 .Enabled = True
 Exit Sub
 End If
 
 If MinutesElapsed Mod 60 = 0 Then
 If MinutesElapsed = 0 Then
 RaiseEvent Timer
 ResetTimer True
 Exit Sub
 End If
 ElseIf .Interval = 1 Or .Interval = m_mSeconds Then
 .Interval = 60000#
 RaiseAlarms
 
 'Reduce minutes remaining by 1
 MinutesElapsed = MinutesElapsed - 1
 
 .Enabled = True
 Exit Sub
 End If
 
 RaiseAlarms
 
 'Reduce minutes remaining by 1
 MinutesElapsed = MinutesElapsed - 1
 
 .Enabled = True
 End With
 
End Sub
Private Sub RaiseAlarms()
 If MinutesElapsed > 60 Then
 'Hours remaining
 RaiseEvent HourAlarm(MinutesElapsed / 60)
 End If
 If MinutesElapsed >= 1 Then
 'Minutes remaining
 RaiseEvent MinuteAlarm(MinutesElapsed)
 End If
End Sub
Private Sub UserControl_InitProperties()
''save default properties
m_TimerMode = TimerMode
m_mSeconds = Interval
m_Minutes = MinuteInterval
m_Hours = HourInterval
m_Enabled = True
 
 
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 'Load property values from storage
 m_TimerMode = PropBag.ReadProperty("TimerMode", 0)
 m_mSeconds = PropBag.ReadProperty("Interval", 0)
 m_Minutes = PropBag.ReadProperty("MinuteInterval", 0)
 m_Hours = PropBag.ReadProperty("HourInterval", 0)
 m_Enabled = PropBag.ReadProperty("Enabled", True)
 
 If (Ambient.UserMode = True) And m_Enabled Then ResetTimer
End Sub
Private Sub UserControl_Resize()
 UserControl.Width = 355
 UserControl.Height = 335
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 
 'Write property values to storage
 Call PropBag.WriteProperty("TimerMode", TimerMode, 0)
 Call PropBag.WriteProperty("Interval", Interval, 0)
 Call PropBag.WriteProperty("MinuteInterval", MinuteInterval, 0)
 Call PropBag.WriteProperty("HourInterval", HourInterval, 0)
 Call PropBag.WriteProperty("Enabled", Enabled, True)
 
End Sub
Private Sub ResetTimer(Optional EventFired As Boolean = False)
 With Timer1
 .Enabled = False 'Stop current timer interval
 .Interval = 1
 msFlag = False
 
 'define hours and minutes into total minutes
 MinutesElapsed = (m_Hours * 60) + m_Minutes
 
 If m_mSeconds = 0 And MinutesElapsed = 0 Then Exit Sub
 
 If EventFired Then
 If TimerMode = OneTime Then
 OneTimeUsed = True
 .Enabled = False
 Else
 OneTimeUsed = False
 .Enabled = True
 End If
 Else
 OneTimeUsed = False
 .Enabled = m_Enabled
 End If
 
 End With
 
End Sub
