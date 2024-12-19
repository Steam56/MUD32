Attribute VB_Name = "Module5"
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10


Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206


Public Const APP_SYSTRAY_ID = 999 'unique identifier

Public Const NOTIFYICON_VERSION = &H3

'Public Type NOTIFYICONDATA
'  cbSize As Long
'  hWnd As Long
'  uID As Long
'  uFlags As Long
'  uCallbackMessage As Long
'  hIcon As Long
'  szTip As String * 64
'End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeout As Integer
  uVersion As Integer
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type
Public Nid As NOTIFYICONDATA

  'uTimeoutAndVersion As Long

Public gNotifyIconData As NOTIFYICONDATA
Public uJarCrypt As JarCrypt

Public Sub GlobalError(ErrVal, uFunc, uModule)
MsgBox ErrVal & Error(ErrVal) & uFunc & uModule
End

End Sub
