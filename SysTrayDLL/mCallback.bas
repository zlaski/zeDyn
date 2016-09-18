Attribute VB_Name = "mCallback"
Option Explicit
Private sTray As SysTray

Public Function Init(SysTray As SysTray)
    Set sTray = SysTray
End Function
Public Function mCallbackFunction(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    mCallbackFunction = sTray.CallBack(hwnd, Msg, wParam, lParam)
End Function


