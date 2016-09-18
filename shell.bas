Attribute VB_Name = "modShell"
Option Explicit

Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Public Enum StartWindowState
    START_HIDDEN = 0
    START_NORMAL = 4
    START_MINIMIZED = 2
    START_MAXIMIZED = 3
End Enum

' app globs
Public gINIFile As String
Public gLogFile As String
Public gCurIP As String
Public gIPDetectURL As String
Public gIPDetectPrefix As String
Public LastState As Integer
Public LngFile As String, RL As Boolean
Public Selitmlst As Integer

Public Function ReadLog()
    On Error GoTo hErr
    If FileExists(gLogFile) = False Then Exit Function
    Open gLogFile For Input As #1
        frmMain.txtLog.Text = Input(LOF(1), 1)
    Close #1
    frmMain.txtLog.SelStart = Len(frmMain.txtLog)
    
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

Public Function WriteLog(Data As String)
    On Error GoTo hErr
        Open gLogFile For Append As #1
            Print #1, Now & vbTab & Data
        Close #1
        frmMain.txtLog.SelStart = Len(frmMain.txtLog)
    
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

Public Function CheckData() As Boolean
    On Error GoTo hErr
    If frmMain.uLogin = "" Or frmMain.uPass = "" Or frmMain.uHosts = "" Then
        If RL = True Then
            MsgBox ReadLng("Messages", "checkinf", LngFile), vbCritical + vbMsgBoxRtlReading + vbMsgBoxRight, App.Title
        Else
            MsgBox ReadLng("Messages", "checkinf", LngFile), vbCritical, App.Title
        End If
        CheckData = False
        frmMain.uStatus.Caption = ReadLng("Messages", "checkinf", LngFile)
        frmMain.uStatus.ForeColor = COLOR_RED
    Else
        CheckData = True
    End If
    
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

Public Function ShellDocument(sDocName As String, _
                    Optional ByVal Action As String = "Open", _
                    Optional ByVal Parameters As String = vbNullString, _
                    Optional ByVal Directory As String = vbNullString, _
                    Optional ByVal WindowState As StartWindowState) As Boolean
    On Error GoTo hErr
    Dim Response
    Response = ShellExecute(&O0, Action, sDocName, Parameters, Directory, WindowState)
    Select Case Response
        Case Is < 33
            ShellDocument = False
        Case Else
            ShellDocument = True
    End Select
    
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

Public Function FileExists(strFile As String) As Boolean
    On Error GoTo hErr

    If PathFileExists(strFile) = 1 Then
        FileExists = True
    ElseIf PathFileExists(strFile) = 0 Then
        FileExists = False
    End If
    
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

Public Function GetINIString(strINIFile As String, strSection As String, strKey As String, _
                                strDefault As String) As String
    On Error GoTo hErr
    Dim strTemp As String * 256     'set string max length to 256 chars
    Dim intLength As Integer
    strTemp = ""
    strTemp = Space$(256)           'initialize string with spaces
    intLength = GetPrivateProfileString(strSection, strKey, strDefault, strTemp, 255, strINIFile)
    GetINIString = Left$(strTemp, intLength)
    
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

Public Sub WriteINIString(strINIFile As String, strSection As String, strKey As String, strvalue As String)
    On Error GoTo hErr
    Dim indx As Integer
    Dim strTemp As String
    
    strTemp = strvalue
    
    'a key value must not contain either a carriage return or a line feed, therefore check for these in
    'the passed string, and substitute a " " if you find one.  This is purely precautionary.
    For indx = 1 To Len(strvalue)
        If Mid$(strvalue, indx, 1) = vbCr Or Mid$(strvalue, indx, 1) = vbLf Then
            Mid$(strvalue, indx) = " "
        End If
    Next indx
    
    indx = WritePrivateProfileString(strSection, strKey, strTemp, strINIFile)
    
Exit Sub
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Sub
End Sub

Public Function ReadLng(Section As String, KeyName As String, FileName As String) As String
    On Error GoTo hErr
    Dim sRet As String
    If FileExists(FileName) = False Then Exit Function
    sRet = String(255, Chr(0))
    ReadLng = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
    
Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function

Public Function ReadLanguage()
    On Error GoTo hErr
'Load language in a *.lng  file
'LngFile = App.Path & "\" & frmMain.Lst.Text & ".lng"
'-------------------------------------------------------------------------------
'Right-Left Or Left-Right
RL = ReadLng("Other", "RightLeft", LngFile)
'----------------------------------------------------------------------------
If RL = True Then
    'Captions
    frmMain.RightToLeft = True
    frmMain.Caption = ReadLng("Main", "Caption", LngFile)
    frmAbout.RightToLeft = True
    '--------------------------------------------------------------------------------
    'Labels
    frmMain.lblUser.Alignment = vbRightJustify
    frmMain.lblUser.RightToLeft = True
    frmMain.lblUser.Caption = ReadLng("Main", "Username", LngFile)
    frmMain.lblPass.Alignment = vbRightJustify
    frmMain.lblPass.RightToLeft = True
    frmMain.lblPass.Caption = ReadLng("Main", "Password", LngFile)
    frmMain.lblStatus.Alignment = vbRightJustify
    frmMain.lblStatus.RightToLeft = True
    frmMain.lblStatus.Caption = ReadLng("Main", "Status", LngFile)
    If frmMain.uStatus.Caption = "" Then
        'frmMain.uStatus.Alignment = vbRightJustify
        frmMain.uStatus.RightToLeft = True
        frmMain.uStatus.Caption = ReadLng("Main", "Starting", LngFile)
        frmMain.uStatus.ForeColor = COLOR_BLUE
    End If
    frmMain.lblHost.Alignment = vbRightJustify
    frmMain.lblHost.RightToLeft = True
    frmMain.lblHost.Caption = ReadLng("Main", "Host", LngFile)
    frmMain.lblInterval.Alignment = vbRightJustify
    frmMain.lblInterval.RightToLeft = True
    frmMain.lblInterval.Caption = ReadLng("Main", "Interval", LngFile)
    frmMain.lblMin.Alignment = vbRightJustify
    frmMain.lblMin.RightToLeft = True
    frmMain.lblMin.Caption = ReadLng("Main", "Minute", LngFile)
    frmMain.lblLog.Alignment = vbRightJustify
    frmMain.lblLog.RightToLeft = True
    frmMain.lblLog.Caption = ReadLng("Main", "Log", LngFile)
    frmMain.lblLang.Alignment = vbRightJustify
    frmMain.lblLang.RightToLeft = True
    frmMain.lblLang.Caption = ReadLng("Main", "Lang", LngFile)
    frmMain.lblCreate.Alignment = vbRightJustify
    frmMain.lblCreate.RightToLeft = True
    frmMain.lblCreate.Caption = ReadLng("Main", "CreateAc", LngFile)
    frmMain.lblFP.Alignment = vbRightJustify
    frmMain.lblFP.RightToLeft = True
    frmMain.lblFP.Caption = ReadLng("Main", "FPass", LngFile)
Else
    'Captions
    frmMain.RightToLeft = False
    frmMain.Caption = ReadLng("Main", "Caption", LngFile)
    frmAbout.RightToLeft = False
    '--------------------------------------------------------------------------------
    'Labels
    frmMain.lblUser.Alignment = vbLeftJustify
    frmMain.lblUser.RightToLeft = False
    frmMain.lblUser.Caption = ReadLng("Main", "Username", LngFile)
    frmMain.lblPass.Alignment = vbLeftJustify
    frmMain.lblPass.RightToLeft = False
    frmMain.lblPass.Caption = ReadLng("Main", "Password", LngFile)
    frmMain.lblStatus.Alignment = vbLeftJustify
    frmMain.lblStatus.RightToLeft = False
    frmMain.lblStatus.Caption = ReadLng("Main", "Status", LngFile)
    If frmMain.uStatus.Caption = "" Then
        'frmMain.uStatus.Alignment = vbLeftJustify
        frmMain.uStatus.RightToLeft = False
        frmMain.uStatus.Caption = ReadLng("Main", "Starting", LngFile)
        frmMain.uStatus.ForeColor = COLOR_BLUE
    End If
    frmMain.lblHost.Alignment = vbLeftJustify
    frmMain.lblHost.RightToLeft = False
    frmMain.lblHost.Caption = ReadLng("Main", "Host", LngFile)
    frmMain.lblInterval.Alignment = vbLeftJustify
    frmMain.lblInterval.RightToLeft = False
    frmMain.lblInterval.Caption = ReadLng("Main", "Interval", LngFile)
    frmMain.lblMin.Alignment = vbLeftJustify
    frmMain.lblMin.RightToLeft = False
    frmMain.lblMin.Caption = ReadLng("Main", "Minute", LngFile)
    frmMain.lblLog.Alignment = vbLeftJustify
    frmMain.lblLog.RightToLeft = False
    frmMain.lblLog.Caption = ReadLng("Main", "Log", LngFile)
    frmMain.lblLang.Alignment = vbLeftJustify
    frmMain.lblLang.RightToLeft = False
    frmMain.lblLang.Caption = ReadLng("Main", "Lang", LngFile)
    frmMain.lblCreate.Alignment = vbLeftJustify
    frmMain.lblCreate.RightToLeft = False
    frmMain.lblCreate.Caption = ReadLng("Main", "CreateAc", LngFile)
    frmMain.lblFP.Alignment = vbLeftJustify
    frmMain.lblFP.RightToLeft = False
    frmMain.lblFP.Caption = ReadLng("Main", "FPass", LngFile)
End If
'Commands buttons
frmMain.uOK.Caption = ReadLng("Main", "OK", LngFile)
frmMain.cmdExit.Caption = ReadLng("Main", "Exit", LngFile)
frmMain.cmdAbout.Caption = ReadLng("Main", "About", LngFile)

If frmMain.Width = 10425 Then
    frmMain.cmdsh.Caption = ReadLng("Main", "LI", LngFile)
Else
    frmMain.cmdsh.Caption = ReadLng("Main", "MI", LngFile)
End If

Exit Function
hErr:
    MsgBox "Error Number:" & vbCrLf & Err.Number & vbCrLf & "Error Description:" & Err.Description
    Exit Function
End Function
