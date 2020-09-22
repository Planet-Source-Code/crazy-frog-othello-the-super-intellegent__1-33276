Attribute VB_Name = "ModuleTray"
Option Explicit

Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type



Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Const WM_USER = &H400
Public Const cbNotify& = WM_USER + 42
Public Const uID& = 61860
Public myNID As NOTIFYICONDATA

Declare Function ShellNotifyIcon Lib "shell32.dll" _
   Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
   lpData As NOTIFYICONDATA) As Long
   
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209


Declare Function CallWindowProc Lib "user32" Alias _
   "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hwnd As Long, ByVal Msg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
   
Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)

Global lpPrevWndProc As Long
Global gHW As Long

Public Sub Hook()
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
      AddressOf WindowProc)
End Sub

Public Sub unhook()
   Dim tmp As Long
   tmp = SetWindowLong(gHW, GWL_WNDPROC, _
      lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
   
   On Error Resume Next
   
'On Error GoTo QuelError
   If wParam = uID Then
      Select Case lParam
         
         Case WM_LBUTTONDOWN
            If MDImain.Visible Then
                MDImain.WindowState = vbNormal
                AppActivate MDImain.Caption
            Else
                Shell FRMoption.TXTshell, 1
            End If
         Case WM_RBUTTONDOWN
            MDImain.Visible = True
            AppActivate MDImain.Caption
      End Select
   End If
   WindowProc = CallWindowProc(lpPrevWndProc, hw, _
      uMsg, wParam, lParam)
'QuelError:
   ' MsgBox Err.Description
End Function


