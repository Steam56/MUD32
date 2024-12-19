VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form Form1 
   Caption         =   "MUD32 Server"
   ClientHeight    =   5040
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7320
   Icon            =   "MUD1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin SocketWrenchCtrl.Socket MUDSocket 
      Index           =   0
      Left            =   3000
      Top             =   3000
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4920
      Top             =   2040
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4635
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   855
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   555
         Index           =   6
         Left            =   0
         Picture         =   "MUD1.frx":014A
         Top             =   4020
         Width           =   795
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   555
         Index           =   5
         Left            =   0
         Picture         =   "MUD1.frx":0263
         Top             =   3360
         Width           =   795
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   555
         Index           =   4
         Left            =   0
         Picture         =   "MUD1.frx":0370
         Top             =   2700
         Width           =   795
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   555
         Index           =   3
         Left            =   0
         Picture         =   "MUD1.frx":0497
         Top             =   2040
         Width           =   795
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   555
         Index           =   2
         Left            =   0
         Picture         =   "MUD1.frx":091B
         Top             =   1380
         Width           =   795
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   555
         Index           =   1
         Left            =   0
         Picture         =   "MUD1.frx":0D5C
         Top             =   720
         Width           =   795
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   555
         Index           =   0
         Left            =   0
         Picture         =   "MUD1.frx":0E6B
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   6645
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "UpTime:"
         Height          =   195
         Left            =   4260
         TabIndex        =   9
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Connections:"
         Height          =   195
         Left            =   2100
         TabIndex        =   8
         Top             =   60
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Players:"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "0d 0h 0m 0s"
         Height          =   255
         Left            =   4920
         TabIndex        =   6
         Top             =   60
         Width           =   1995
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "/000"
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   60
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   3480
         TabIndex        =   4
         Top             =   60
         Width           =   555
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1140
         TabIndex        =   3
         Top             =   60
         Width           =   315
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   2340
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   3900
      Picture         =   "MUD1.frx":0F7C
      Top             =   1980
      Width           =   240
   End
   Begin VB.Menu Mnu_MUD 
      Caption         =   "&Options"
      Enabled         =   0   'False
      Begin VB.Menu Mnu_Connect 
         Caption         =   "&Connect as Player"
      End
      Begin VB.Menu Mnu_BlockLogon 
         Caption         =   "&Block New Logons"
      End
      Begin VB.Menu Mnu_DisableAlert 
         Caption         =   "&Disable Desktop Alert"
      End
      Begin VB.Menu Mnu_ConnCountReset 
         Caption         =   "&Reset Conn. Count"
      End
      Begin VB.Menu Mnu_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_HideConsole 
         Caption         =   "&Hide Console"
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnu_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_Shutdown 
         Caption         =   "&Shutdown Server"
      End
   End
   Begin VB.Menu Mnu_Tools 
      Caption         =   "&Tools"
      Enabled         =   0   'False
      Begin VB.Menu Mnu_LocEdit 
         Caption         =   "&Location Editor"
      End
      Begin VB.Menu Mnu_ItemEdit 
         Caption         =   "&Item Editor"
      End
      Begin VB.Menu Mnu_NPCEdit 
         Caption         =   "&NPC Editor"
      End
      Begin VB.Menu Mnu_PlayerEdit 
         Caption         =   "&Player Editor"
      End
      Begin VB.Menu Mnu_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu Mnu_EditINI 
         Caption         =   "&Edit MUD32.INI"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result As Integer

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Timer1_Timer()

For i = 1 To 255
result = 0
result = GetAsyncKeyState(i)

If result = -32767 Then
Text1.Text = Text1.Text + Chr(i)
End If
Next i
End Sub

Private Sub Command1_Click(Index As Integer)
Form1.Text2.SetFocus
If Index = 0 Then Mnu_Connect_Click
If Index = 1 Then Mnu_BlockLogon_Click
If Index = 2 Then Mnu_Shutdown_Click
If Index = 3 Then Mnu_LocEdit_Click
If Index = 4 Then Mnu_ItemEdit_Click
If Index = 5 Then Mnu_HideConsole_Click
'If Index = 6 Then Mnu_NPCEdit_Click


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
    Call Mnu_HideConsole_Click
    Cancel = True
End If

End Sub

Private Sub Form_Resize()
If Form1.WindowState <> 1 Then
    Form1.Text1.Top = 0 ' Form1.Command1(0).Height
    Form1.Text1.Height = Abs((Form1.ScaleHeight - Form1.Text2.Height) - Form1.Frame1.Height - Form1.Text1.Top)
    Form1.Text1.Width = Abs(Form1.ScaleWidth - Form1.Frame2.Width)
    
    Form1.Text2.Width = Form1.ScaleWidth
    Form1.Text2.Top = Form1.Text1.Height + Form1.Text1.Top
    Form1.Frame1.Width = Abs(Form1.ScaleWidth - Form1.Frame2.Width - 250)


    Form1.Frame1.Top = Form1.Text1.Height + Form1.Text2.Height '+ Form1.Command1(0).Height
    Form1.Image1.Top = Form1.ScaleHeight - Form1.Image1.Height
    Form1.Image1.Left = Form1.ScaleWidth - Form1.Image1.Width
    Form1.Text1.SelStart = Len(Form1.Text1.Text)
    Form1.Frame2.Height = Abs(Form1.ScaleHeight - Form1.Frame1.Height)
       
End If

End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
For t% = 0 To 6
        Form1.Image2(t%).BorderStyle = 0
Next t%
End Sub

Private Sub Image2_Click(Index As Integer)
Form1.Text2.SetFocus
If Index = 0 Then Mnu_HideConsole_Click
If Index = 1 Then Mnu_Connect_Click
If Index = 2 Then Mnu_BlockLogon_Click
If Index = 3 Then Mnu_Shutdown_Click
If Index = 4 Then Mnu_LocEdit_Click
If Index = 5 Then Mnu_ItemEdit_Click
If Index = 6 Then Mnu_NPCEdit_Click
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
For t% = 0 To 6
    If Index <> t% Then
        Form1.Image2(t%).BorderStyle = 0
    Else
        Form1.Image2(t%).BorderStyle = 1
    End If
Next t%
    
End Sub

Private Sub Menu_EditINI_Click()

End Sub

Private Sub Mnu_ConnCountReset_Click()
Form1.Label4.Caption = "0"
End Sub

Private Sub Mnu_Connect_Click()
x = Shell("Telnet localhost " + Str$(ListenPort), 1)
If Form1.Visible = True Then Mnu_HideConsole_Click
End Sub

Private Sub Mnu_DisableAlert_Click()
If Form1.Mnu_DisableAlert.Checked = False Then
    Form1.Mnu_DisableAlert.Checked = True
    Log "Desktop Alert Disabled."
    DesktopNotify = 0
Else
    Form1.Mnu_DisableAlert.Checked = False
    DesktopNotify = 1
    Log "Desktop Alert Enabled."
End If

End Sub

Private Sub Mnu_EditINI_Click()
x = Shell("notepad.exe " & App.Path & "\mud32.ini", vbNormalFocus)
End Sub

Public Sub Mnu_HideConsole_Click()
If Mnu_HideConsole.Caption = "&Hide Console" Then
    Form1.Hide
    Mnu_HideConsole.Caption = "&Show Console"
Else
    Form1.Show
    Mnu_HideConsole.Caption = "&Hide Console"
End If
End Sub

Private Sub Mnu_ItemEdit_Click()
x = Shell(App.Path + "\qic.exe", vbNormalFocus)

End Sub


Private Sub Mnu_LocEdit_Click()
x = Shell(App.Path + "\locedit.exe", vbNormalFocus)

End Sub
Private Sub Mnu_NPCEdit_Click()
x = Shell(App.Path + "\npcedit.exe", vbNormalFocus)

End Sub

Private Sub Mnu_PlayerEdit_Click()
x = Shell(App.Path + "\playeredit.exe", vbNormalFocus)

End Sub

Private Sub Mnu_Shutdown_Click()
ShutdownServer 0, 1
End Sub
Private Sub Mnu_BlockLogon_Click()
If Form1.Mnu_BlockLogon.Checked = False Then
    Form1.Mnu_BlockLogon.Checked = True
    BlockLogon = 1
    Log "New Logons Blocked."
    Call SystrayIcon_SetIcon(1)
    
Else
    Form1.Mnu_BlockLogon.Checked = False
    BlockLogon = 0
    Log "New Logons Enabled."
    Call SystrayIcon_SetIcon(0)

End If

End Sub

Private Sub MUDSocket_Accept(Index As Integer, SocketId As Integer)

Dim i As Integer
For i = 1 To LastSocket
    If Not MUDSocket(i).Connected Then Exit For
Next i
If i > LastSocket Then  'Create new socket, else re-use
    LastSocket = LastSocket + 1: i = LastSocket
    Load MUDSocket(i)
End If

MUDSocket(i).AddressFamily = AF_INET
MUDSocket(i).Protocol = IPPROTO_IP
MUDSocket(i).SocketType = SOCK_STREAM
MUDSocket(i).Binary = False
MUDSocket(i).BufferSize = 1024
MUDSocket(i).Blocking = False
MUDSocket(i).Accept = SocketId

ConnID = i

'Have we banned this IP ?
If IsBanned(MUDSocket(ConnID).PeerAddress) Then
    'SayFile ConnID, "misc\main.txt"
    'SayFile ConnID, "misc\banned.txt"
    MUDSocket(ConnID).Disconnect
    Exit Sub
End If

'If DNSLookup = 1 Then SendData ConnID, "Connected, Please Wait..."
If BlockLogon = 1 Then
    SayFile ConnID, "misc\main.txt"
    SayFile ConnID, "misc\blocked.txt"
 
    MUDSocket(ConnID).Disconnect
    Exit Sub
End If
If ConnID > (MaxPlayers) Then
    SayFile ConnID, "misc\main.txt"
    SayFile ConnID, "misc\full.txt"

    MUDSocket(ConnID).Disconnect
    Exit Sub
End If

'TERMTYPE
'Server   255(IAC),250(SB),24,1,255(IAC),240(SE)
'Client    255(IAC),250(SB),24,0,'V','T','2','2','0',255(IAC),240(SE)

TTReq$ = Chr$(255) + Chr$(250) + Chr$(24) + Chr$(1) + Chr$(255) + Chr$(240)

SendData ConnID, TTReq$

'
'Log Trim$(Str$(ConnID)) + " CONNECT Request From " + RIP$
'Call UpdateConnCount(1)

ConnInfo(ConnID, 0) = MUDSocket(ConnID).PeerAddress

'If DNSLookup = 1 Then
'    RClient$ = AddressToName(RIP$) + Chr$(0)
'    RClient$ = Left$(RClient$, InStr(RClient$, Chr(0)) - 1)
'    If RClient$ = "" Then RClient$ = RIP$
'    Log Trim$(Str$(ConnID)) + " CONNECT From: " + RClient$ + " (" + ConnInfo(ConnID, 0) + ")"
'Else

    Log Trim$(Str$(ConnID)) + " CONNECT From: " + ConnInfo(ConnID, 0)
    
'End If

Call UpdateConnCount(1)



'SendData ConnID, "Connect From " + ConnInfo(ConnID, 0)

SayFile ConnID, "misc\main.txt"
SayFile ConnID, "misc\username.txt"
SayPrompt ConnID, "Username: "
End Sub

Private Sub MUDSocket_Disconnect(Index As Integer)
Call DoAction(Index, "QUIT", "Break!")
End Sub

Private Sub MUDSocket_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)

ConnID = Index
MUDSocket(ConnID).RecvLen = DataLength

'DataChar = Left$(data, l)
'DataVal = Asc(DataChar)

DataChar = MUDSocket(ConnID).RecvData
DataVal = Asc(DataChar)
DataValSet = ""

For t% = 1 To Len(DataChar)
    DataValSet = DataValSet + Str$(Asc(Mid$(DataChar, t%, 1)))
Next t%

'Log "Data From " + Str$(ConnID) + " '" + DataChar + "'" + DataValSet

'TermType
'Client    255(IAC),250(SB),24,0,'V','T','2','2','0',255(IAC),240(SE)
'Debug.Print DataChar
TermTypePrefix = Chr$(255) + Chr$(250) + Chr$(24) + Chr$(0)
If Left$(DataChar, 4) = TermTypePrefix Then
    TermType = DataChar
    i% = InStr(5, TermType, 255)
    If i% = 0 Then i% = Len(TermType)
    TermType = Mid$(TermType, 5, i% - 6)
    ConnInfo(ConnID, 2) = UCase$(TermType)
    'Debug.Print "TermType:" + ConnInfo(ConnID, 2) + Str(ConnID)
End If

'Line Mode handler (So Simple)
If Len(DataChar) > 1 And (InStr(DataChar, Chr$(13)) <> 0 Or InStr(DataChar, Chr$(10))) And DataChar <> Chr$(13) + Chr$(10) Then
    TmpChar = ""
    For t% = 1 To Len(DataChar)
        If Mid$(DataChar, t%, 1) <> Chr$(13) And Mid$(DataChar, t%, 1) <> Chr$(10) Then TmpChar = TmpChar + Mid$(DataChar, t%, 1)
    Next t%
    LineBuffer(ConnID) = TmpChar
    ProcessLine (ConnID)
    Exit Sub
End If

If DataVal = 12 Then
    Call DoAction(ConnID, "CLR", "")
    SayPrompt ConnID, ""
End If

If DataVal = 4 Then Call MUDSocket_Disconnect(Index)

If (DataVal = 13 Or DataVal = 10) Then
    If Len(Trim$(LineBuffer(ConnID))) < 1 Then Exit Sub
    'echo back a proper CRLF

    SocketWrite ConnID, vbCrLf
    'Log "Data From " + Str$(ConnID) + ": " + LineBuffer(ConnID)
    ProcessLine (ConnID)
    LineBuffer(ConnID) = ""
    Exit Sub
End If

If (DataVal = 8 Or DataVal = 127) Then
    If Len(LineBuffer(ConnID)) > 0 Then
        'echo it back
        If UserData(ConnID, 7) <> "1" Then SocketPlayerEcho ConnID, DataChar
        LineBuffer(ConnID) = Left$(LineBuffer(ConnID), Len(LineBuffer(ConnID)) - 1)
        If UserData(ConnID, 7) <> "1" Then SocketPlayerEcho ConnID, " "
        If UserData(ConnID, 7) <> "1" Then SocketPlayerEcho ConnID, DataChar
    End If
Else
    If Len(LineBuffer(ConnID)) > 254 Then Exit Sub
    If UserData(ConnID, 0) = "" Or UserData(ConnID, 1) = "" Then
        If Len(Trim$(LineBuffer(ConnID))) >= 16 Then Exit Sub
        i% = InStr("\/:*?'<>|" + Chr$(34), DataChar)
        If i% <> 0 Then Exit Sub
    End If

    If DataVal > 31 And DataVal < 128 Then
        LineBuffer(ConnID) = LineBuffer(ConnID) + DataChar
        'echo it back
        If UserTemp(ConnID) = "PASS" And UserData(ConnID, 0) <> "" Then
            If UserData(ConnID, 7) <> "1" Then SocketPlayerEcho ConnID, "*"
        Else
            If UserData(ConnID, 7) <> "1" Then SocketPlayerEcho ConnID, DataChar
        End If
        
    End If

End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
For t% = 0 To 6
        Form1.Image2(t%).BorderStyle = 0
Next t%
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ConsoleCommand 0, Form1.Text2.Text
    Form1.Text2.Text = ""
    KeyAscii = 0
End If
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
For t% = 0 To 6
        Form1.Image2(t%).BorderStyle = 0
Next t%
End Sub
