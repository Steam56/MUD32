VERSION 5.00
Begin VB.Form frmNotify 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1260
   ClientLeft      =   2505
   ClientTop       =   2460
   ClientWidth     =   1560
   ControlBox      =   0   'False
   Icon            =   "MUD_Notify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   Begin VB.Timer notifyTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   60
      Top             =   765
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "dTitle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "MUD_Notify.frx":0CFA
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "dMessage"
      Height          =   630
      Left            =   180
      TabIndex        =   0
      Top             =   510
      Width           =   1215
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1260
      Left            =   0
      Picture         =   "MUD_Notify.frx":0E44
      Top             =   0
      Width           =   1560
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type APPBARDATA
        cbSize As Long
        hWnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long '  message specific
End Type

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Const ABM_GETTASKBARPOS = &H5

Private Declare Function sndPlaySoundA Lib "WINMM.DLL" (ByVal lpszSoundName As String, ByVal ValueFlags As Long) As Long
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 '  Maintenance string for PSS usage
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

' dwPlatforID Constants
Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2
'-- End --'

Private Sub Form_Load()
'---------------------------------------------------------
'NotifyDisplay "MUD32", "Yhama-Llama has joined the world"
'---------------------------------------------------------
End Sub

Public Sub NotifyDisplay(dTitle$, dMessage$)

frmNotify.Label1.Caption = dMessage$
frmNotify.Label3.Caption = dTitle$
notifyTimer.Enabled = True

End Sub

Private Sub Label1_Click()
frmNotify.Hide
notifyTimer.Enabled = False
End Sub



Private Sub notifyTimer_Timer()
Static rMode As Integer
DesktopWorkArea xOffset, yOffset, xMax, yMax

If rMode = 0 Then
    PlayNotify
    frmNotify.Left = xMax - frmNotify.Width
    frmNotify.Top = yMax
    frmNotify.Show
    rMode = 1
End If

If rMode > 1 And rMode < 40 Then
    rMode = rMode + 1
    Exit Sub
End If

MinTop = yMax - frmNotify.Height
Stepper = (yMax - MinTop) / 10

If rMode = 1 Then
    CurPos = frmNotify.Top
    If CurPos > MinTop Then
        NewPos = frmNotify.Top - Stepper
        If NewPos < MinTop Then NewPos = MinTop
        frmNotify.Top = NewPos
        Exit Sub
    Else
        rMode = 2
        Exit Sub
    End If
Else
    CurPos = frmNotify.Top
    If CurPos < yMax Then
        frmNotify.Top = frmNotify.Top + Stepper
        Exit Sub
    Else
        rMode = 0
        notifyTimer.Enabled = False
        frmNotify.Hide
        Exit Sub
    End If
End If

End Sub

Public Sub DesktopWorkArea(xOffset, yOffset, xMax, yMax)
Call GetTaskbarInfo(uTop, uBottom, uLeft, uRight)

xMax = Screen.Width
yMax = Screen.Height

If uTop = -2 And uLeft = -2 Then
    If uRight < uBottom Then
        xOffset = uRight * 15
    Else
        yOffset = uBottom * 15
    End If
ElseIf uTop = -2 And uLeft > 0 Then xMax = uLeft * 15
ElseIf uTop > 0 And uLeft = -2 Then yMax = uTop * 15
End If

End Sub

Sub GetTaskbarInfo(xTop, xBottom, xLeft, xRight)
Dim abData As APPBARDATA
Dim lResult As Long

    lResult = SHAppBarMessage(ABM_GETTASKBARPOS, abData)
    xTop = abData.rc.Top
    xBottom = abData.rc.Bottom
    xLeft = abData.rc.Left
    xRight = abData.rc.Right
    'Debug.Print xTop, xBottom, xLeft, xRight
    
End Sub

Public Sub PlayNotify()
  Dim Ret As Long
  'upath = Environ("systemroot")
  'Ret = sndPlaySoundA(upath + "\media\notify.wav", SND_ASYNC Or SND_NODEFAULT)
  Ret = sndPlaySoundA("notify.wav", SND_ASYNC Or SND_NODEFAULT)
End Sub

Public Function GetWindowsVersion(VerType)

Dim tOSVer As OSVERSIONINFO
    
   ' First set length of OSVERSIONINFO
   ' structure size
   tOSVer.dwOSVersionInfoSize = Len(tOSVer)
   ' Get version information
   GetVersionEx tOSVer
   ' Determine OS type
   With tOSVer
      
      Select Case .dwPlatformId
         Case VER_PLATFORM_WIN32_NT
            ' This is an NT version (NT/2000)
            ' If dwMajorVersion >= 5 then
            ' the OS is Win2000
            If .dwMajorVersion >= 5 Then
                If .dwMinorVersion > 0 Then
                    MajorOS = "Windows XP"
                Else
                    MajorOS = "Windows 2000"
                End If
            Else
               MajorOS = "Windows NT"
            End If
         Case Else
            ' This is Windows 95/98/ME
            If .dwMajorVersion >= 5 Then
               MajorOS = "Windows ME"
            ElseIf .dwMajorVersion = 4 And .dwMinorVersion > 0 Then
               MajorOS = "Windows 98"
            Else
               MajorOS = "Windows 95"
            End If
         End Select
         ' Check for service pack
         ServicePack = Left(.szCSDVersion, InStr(1, .szCSDVersion, Chr$(0)) - 1)
         ' Get OS version
         NumericVersion = .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber
        
    End With

    Select Case VerType
        Case 2
            GetWindowsVersion = NumericVersion
        Case 1
            If ServicePack <> "" Then
                GetWindowsVersion = " SP" + Right$(ServicePack, 1)
            Else
                GetWindowsVersion = ""
            End If
        Case Else
            GetWindowsVersion = "Microsoft " & MajorOS
    End Select
    
End Function


