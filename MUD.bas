Attribute VB_Name = "Module1"
Global LineBuffer(64) As String
Global ConnInfo(64, 10) As String
Global UserData(64, 50) As String
Global PlayerLocHistory(64) As String * 1500
Global UserTemp(64) As String
Global Levels(20, 4) As String
Global MaxLevels As Integer
Global PlayerExit(100, 10, 3) As String
Global TotalPlayerExit(100) As Integer
Global ItemArray(1000, 7) As String
Global ItemData(64, 9) As String
Global ListenPort As Integer
Global ScoreLoc As Integer
Global StartLoc As Integer
Global NoticeBoardLoc As Integer
Global MinRespawn As Integer
Global MaxRespawn As Integer
Global MaxPlayers As Integer
Global StickyQuit As Integer
Global DNSLookup As Integer
Global SystemTray As Integer
Global RespawnList(1000, 1) As Integer
Global DDEReporting As Integer
Global DDETopic As String
Global DDEItem As String
Global DDEMsgPrefix As String
Global GodLevel As Integer
Global BalanceFights As Integer
Global IdleProtect As Integer
Global DesktopNotify As Integer
Global TopTenArray(10, 3)
Global LocFlags(64) As String
Global VTBold(1, 1) As String
Global AnsiTable(1, 20) As String
Global AnsiCodes(30, 1) As String
Global WebRoot As String
Global SuperUser As String
Global LogFile As String
Global BuildVersion As String
Global WindowsVersion As String
Global BlockLogon As Integer
Global MudStartTime As Date
Global LastSocket As Integer
Global ShowExits As Integer
Sub Log(Text$)
Text$ = "[" + Str$(Date) + " " + Time$ + "] " + Text$
Form1.Text1.SelStart = Len(Form1.Text1.Text)
Form1.Text1.Text = Form1.Text1.Text + Text$ + Chr$(13) + Chr$(10)
If Len(Form1.Text1.Text) > 8000 Then Form1.Text1.Text = Right$(Form1.Text1.Text, 8000)
Form1.Text1.SelStart = Len(Form1.Text1.Text)
Form1.Text1.Refresh
If Form1.Visible = True Then Form1.Text2.SetFocus
myHandle = FreeFile
If LogFile <> "" Then
    Open App.Path + "\" + LogFile For Append Access Write As myHandle
    Print #myHandle, Text$
    Close myHandle
End If
End Sub
Sub DesktopAlert(dTitle$, dMessage$)
If DesktopNotify <> 0 Then

    If InStr(WindowsVersion, "XP") <> 0 Or InStr(WindowsVersion, "2000") <> 0 Then
        'balloon tip
        With Nid
            .cbSize = Len(Nid)
            .hWnd = Form3.hWnd
            .uID = vbNull
            .uFlags = NIF_INFO
            .dwInfoFlags = 1 'nIconIndex
            .szInfoTitle = Form1.Caption & vbNullChar
            .szInfo = dMessage$ & vbNullChar
            .uTimeout = 15
        End With
        
        Shell_NotifyIcon NIM_MODIFY, Nid
        'balloon tip end
    Else
        frmNotify.NotifyDisplay dTitle$, dMessage$
    End If

End If
End Sub
Sub UpdateLastOn(PlayerName$)
    Dim LastOn(20) As String
    Counter = 0
    Handle = FreeFile
    If IfExist(App.Path + "\misc\laston.dat") Then
        Open App.Path + "\misc\laston.dat" For Input Access Read As #Handle
        Do
            Line Input #Handle, Ln$
            LastOn(Counter) = Ln$
            Counter = Counter + 1
        Loop Until EOF(Handle)
        Close #Handle
    End If
    If Counter = 20 Then Offset = 18 Else Offset = 19
    Handle = FreeFile
    Open App.Path + "\misc\laston.dat" For Output Access Write As #Handle
    Print #Handle, PlayerName$ & " on " & Format(Date, "dd mmmm yyyy") & " at " & Format(Time, "hh:mm")
    For t% = 0 To Offset
        If LastOn(t%) <> "" Then Print #Handle, LastOn(t%) Else t% = 21
    Next t%
    Close #Handle
End Sub
Sub SysTrayIcon_Create()
'load From
Load Form3

'Init Icon
With gNotifyIconData
  .cbSize = Len(gNotifyIconData)
  .hWnd = Form3.hWnd
  .uID = vbNull
  .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  .uCallbackMessage = WM_MOUSEMOVE
  .hIcon = Form3.Icon
  .szTip = Form1.Caption & vbNullChar
End With

'Create Icon
Shell_NotifyIcon NIM_ADD, gNotifyIconData



End Sub
Sub SystrayIcon_SetIcon(IconNo)
If IconNo = 0 Then
    gNotifyIconData.hIcon = Form3.Icon
    Shell_NotifyIcon NIM_MODIFY, gNotifyIconData
Else
    gNotifyIconData.hIcon = Form2.Icon
    Shell_NotifyIcon NIM_MODIFY, gNotifyIconData
End If

End Sub
Sub SystrayIcon_SetTip(Tip$)
gNotifyIconData.szTip = Tip$ + vbNullChar
Shell_NotifyIcon NIM_MODIFY, gNotifyIconData
End Sub

Sub SysTrayIcon_Delete()
'Delete Icon
Shell_NotifyIcon NIM_DELETE, gNotifyIconData
Unload Form3

End Sub

Public Sub Main()
'On Error GoTo MainErrorTrap
Randomize Timer
MudStartTime = Now
ChDir App.Path
LogFile = ReadINI("mud32.ini", "LOGFILE")
BuildVersion = Trim$(Str$(App.Major)) + "." + Right$("00" + Trim$(Str$(App.Minor)), 2) + "." + Right$("0000" + Trim$(Str$(App.Revision)), 4)
WindowsVersion = frmNotify.GetWindowsVersion(0) + frmNotify.GetWindowsVersion(1)  ' + " (" + frmNotify.GetWindowsVersion(2) + ")"

SilentGUI = Val(ReadINI("mud32.ini", "SILENT"))
If SilentGUI <> 1 Then
    Form1.Show
    Form1.Refresh
End If

Log "MUD Server Build " + BuildVersion

Log "Server Start..."

ListenPort = Val(Command$)
If ListenPort = 0 Then ListenPort = 23

Log "Loading Defaults from MUD32.INI"
Amsg$ = ReadINI("mud32.ini", "ADMINMSG") + ""
Log Amsg$
ListenPort = Val(ReadINI("mud32.ini", "DEFAULTPORT"))
StartLoc = Val(ReadINI("mud32.ini", "STARTLOC"))
ScoreLoc = Val(ReadINI("mud32.ini", "SCORELOC"))
ShowExits = Val(ReadINI("mud32.ini", "SHOWEXITS"))
NoticeBoardLoc = Val(ReadINI("mud32.ini", "NOTICEBOARDLOC"))
MinRespawn = Val(ReadINI("mud32.ini", "MINRESPAWNTIME"))
MaxRespawn = Val(ReadINI("mud32.ini", "MAXRESPAWNTIME"))
MaxPlayers = Val(ReadINI("mud32.ini", "MAXPLAYERS"))
Form1.Caption = ReadINI("mud32.ini", "WINTITLE") + " - " + App.Title + " (" + BuildVersion + ")"
GodLevel = Val(ReadINI("mud32.ini", "GODLEVEL"))
SuperUser = ReadINI("mud32.ini", "SUPERUSER")
StickyQuit = ReadINI("mud32.ini", "STICKYQUIT")
DNSLookup = Val(ReadINI("mud32.ini", "DNSLOOKUP"))

If SilentGUI <> 1 Then
    DesktopNotify = Val(ReadINI("mud32.ini", "DesktopNotify"))
    RunHidden = Val(ReadINI("mud32.ini", "RUNHIDDEN"))
    SystemTray = Val(ReadINI("mud32.ini", "SYSTEMTRAY"))
Else
    Log "Running MUD32 in Silent Mode; No GUI, No Systray, No Balloon Tips"
End If

If DesktopNotify = 0 Then Form1.Mnu_DisableAlert.Checked = True
WebRoot = ReadINI("mud32.ini", "WEBROOT")
If WebRoot = "" Then WebRoot = App.Path
If Right$(WebRoot, 1) = "\" And Mid$(WebRoot, Len(WebRoot) - 2, 1) <> ":" Then WebRoot = Left$(WebRoot, Len(WebRoot) - 1)
If SystemTray = 1 Then
    Call SysTrayIcon_Create
    Log "Creating System Tray Icon."
    Form1.Mnu_HideConsole.Enabled = True
End If
DDEReporting = Val(ReadINI("mud32.ini", "DDEREPORTING"))
If DDEReporting = 1 Then
    DDETopic = ReadINI("mud32.ini", "DDETOPIC")
    DDEItem = ReadINI("mud32.ini", "DDEITEM")
    DDEMsgPrefix = ReadINI("mud32.ini", "DDEMSGPREFIX")
    Log "DDE Reporting is ON."
End If

If IfExist(App.Path + "\misc\top10.lst") <> True Then
    Log "Creating New Top10 File"
    Call CreateTopTenFile
End If

Log "Enabling VT100 Codes"
Call SetupVT100Codes

Log "Enabling ANSI Codes"
Call SetupAnsiCodes("misc\ansi.ini")

Log "Building help\commands.txt"
Call HelpCommands

Log "Please Wait, Rebuilding Location Cache"
Call RebuildUsedLocCache

BalanceFights = Val(ReadINI("mud32.ini", "BALANCEFIGHTS"))
IdleProtect = Val(ReadINI("mud32.ini", "IDLEPROTECT"))

LoadLevelArray
LoadItemArray
NPC_Load

For t% = 0 To 1000
    RespawnList(t%, 0) = -1
Next t%

If SuperUser <> "" Then
    Log "Super User is " + SuperUser
Else
    Log "*** WARNING: THERE IS NO SUPERUSER, YOU CANNOT ADMIN THE WORLD WITHOUT ONE, ADD 'SuperUser=<Your Player Name>' to MUD32.INI"
End If

Log "Registering JarCrypt.DLL..."
x = Shell("regsvr32 /s " & Chr(34) & App.Path & "\JarCrypt.dll" & Chr(34))

StartListening ListenPort
Form1.Label5.Caption = "/" + Trim$(Str$(MaxPlayers))
Log "Ready For Players. :-)"
Form1.Mnu_MUD.Enabled = True
Form1.Mnu_Tools.Enabled = True
If RunHidden = 1 Then
    Form1.Mnu_HideConsole.Caption = "&Show Console"
    Form1.Hide
End If

DesktopAlert "MUD32", "Has started successfully, and players can now log on."
Exit Sub

MainErrorTrap:
'Log "ERROR: " & Err & " (" & Error(Err) & ")"
'Log "MUD32 is Restarting..."
'End

End Sub

Sub StartListening(Port)


Form1.MUDSocket(0).AddressFamily = AF_INET
Form1.MUDSocket(0).Protocol = IPPROTO_IP
Form1.MUDSocket(0).SocketType = SOCK_STREAM
Form1.MUDSocket(0).Blocking = False
Form1.MUDSocket(0).LocalPort = Port
x = Form1.MUDSocket(0).Listen

Log "Listen Started on Port" + Str$(Port) + " (Return Code: " + Str$(x) + ")"

If x <> 0 Then
    Log "ERROR: MUD Server Did Not Start."
    MsgBox "MUD Server Did Not Start", 16, "Error"
    End
End If

LastSocket = 0


End Sub


Sub SendData(ID, Txt$)
TermWidth = Val(UserData(ID, 22))
If TermWidth = 0 Then TermWidth = 75

Text$ = Txt$

'InsertColor ID, Text$, DummyValue
SocketWrite ID, vbCr

If Len(Text$) > TermWidth Then ' it's too big for one line of text
    Offset = 1
    Offset2 = 1
    Do
        j% = i%
        If j% = 0 Then j% = 1
        i% = InStr(Offset, Text$, " ")
        If i% = 0 Then Exit Do
        
        If i% - Offset2 > TermWidth Then
            Txt$ = Mid$(Text$, Offset2, j% - Offset2) + vbCrLf
            InsertColor ID, Txt$, DummyValue
            SocketWrite ID, Txt$
            Offset2 = j% + 1
            
        End If
        Offset = j% + 1
    Loop
    If Offset2 < Len(Text$) And i% = 0 Then
        Txt$ = Mid$(Text$, Offset2) + vbCrLf
        InsertColor ID, Txt$, DummyValue
        SocketWrite ID, Txt$
    End If
Else 'the whole line is smaller than termwidth
    Text$ = Left$(Text$ + Space$(TermWidth), TermWidth) + vbCrLf
    InsertColor ID, Text$, DummyValue
    SocketWrite ID, Text$
End If

End Sub
Sub SocketWrite(ID, Txt)
Z$ = Txt
RetVal = Form1.MUDSocket(ID).Write(Z$, Len(Z$))
End Sub
Sub SocketPlayerEcho(ID, Txt)
Z$ = Txt
RetVal = Form1.MUDSocket(ID).Write(Z$, Len(Z$))
End Sub
Sub SocketClose(ID)
Form1.MUDSocket(ID).Disconnect
End Sub

Sub ProcessLine(ID)
ResetANSIColor ID

'LineToSend$ = "You Typed:" + LineBuffer(ID)
'SendData ID, LineToSend$

' Process Input
Ln$ = LineBuffer(ID) + " "
LineBuffer(ID) = ""

BadChars$ = "\/:*?'<>|" + Chr$(34)
BadList$ = ""
If ConnInfo(ID, 1) = "99" Then
    If Left$(UCase$(Ln$), 4) <> "SAY " And Left$(UCase$(Ln$), 3) <> "ME " And Left$(UCase$(Ln$), 7) <> "PROFILE " Then
        For t% = 1 To Len(BadChars$)
            i% = InStr(Ln$, Mid$(BadChars$, t%, 1))
            If i% <> 0 Then BadList$ = BadList$ + Mid$(BadChars$, t%, 1)
        Next t%
        If BadList$ <> "" Then
            SendData ID, "Sorry, use of " + BadList$ + " is not allowed with this command."
            SayPrompt ID, ""
            Exit Sub
        End If
    End If
End If


Verb$ = Trim$(Left$(Ln$, InStr(Ln$, " ")))
Param$ = Trim$(Mid$(Ln$, InStr(Ln$, " ")))

Ln$ = Trim$(Ln$)

If ConnInfo(ID, 1) = "1" Then
    If PlayerExist(Ln$) <> 0 Or Left$(UCase$(Ln$), Len("GUEST")) = "GUEST" Then
        SendData ID, "Sorry, That username is already taken, please try again."
        SayPrompt ID, "Your New Username: "
        Exit Sub
    End If
    For t% = 1 To Len(BadChars$) + 1
        i% = InStr(Ln$, Mid$(BadChars$ + " ", t%, 1))
        If i% <> 0 Then BadList$ = BadList$ + Mid$(BadChars$ + " ", t%, 1)
    Next t%
    If BadList$ <> "" Then
        SendData ID, "Sorry, illegal characters in username (" + BadList$ + "), please try again."
        SayPrompt ID, "Your New Username: "
        Exit Sub
    End If
    If InvalidName(Ln$) = True Then
        SendData ID, "Sorry, that username is disallowed, please try again." '+ BadList$
        SayPrompt ID, "Your New Username: "
        Exit Sub
    End If
        
    
    ConnInfo(ID, 1) = "2"
    UserData(ID, 0) = Ln$
    SayPrompt ID, "Your New Password: "
    Exit Sub
End If

If ConnInfo(ID, 1) = "2" Then
    For t% = 1 To Len(BadChars$) + 1
        i% = InStr(Ln$, Mid$(BadChars$ + " ", t%, 1))
        If i% <> 0 Then BadList$ = BadList$ + Mid$(BadChars$ + " ", t%, 1)
    Next t%
    If BadList$ <> "" Then
        SendData ID, "Sorry, illegal characters in password. Please Try Again."
        SayPrompt ID, "Your New Password: "
        Exit Sub
    End If
    ConnInfo(ID, 1) = "3"
    UserData(ID, 1) = Crypt(Ln$)
    SayPrompt ID, "Your Gender (M/F): "

    Exit Sub
End If

If ConnInfo(ID, 1) = "3" Then
    x$ = UCase$(Left$(Verb$, 1))
    If x$ <> "M" And x$ <> "F" Then
        SayPrompt ID, "Your Gender (M/F): "
        Exit Sub
    End If
    ConnInfo(ID, 1) = "4"
    UserData(ID, 2) = x$
    If ConnInfo(ID, 2) = "ANSI" Or ConnInfo(ID, 2) = "VT100" Then
        SendData ID, "Your software supports colour."
    Else
         SendData ID, "Your software MAY NOT support colour."
    End If
    SayPrompt ID, "Do you want colour (Y/N) ? "

    Exit Sub
End If

If ConnInfo(ID, 1) = "4" Then
    x$ = UCase$(Left$(Verb$, 1))
    If x$ <> "Y" And x$ <> "N" Then

        SayPrompt ID, "Do you want colour (Y/N) ? "

        Exit Sub
    End If
    If x$ = "Y" Then UserData(ID, 21) = "1" Else UserData(ID, 21) = "0"
    UserData(ID, 3) = "0"
    
    ConnInfo(ID, 1) = ""
    SaveUser ID
    SendData ID, "Account Created... Welcome " + UserData(ID, 0) + "!"
    
    Call PutPlayerInGame(ID)
    Call DoAction(ID, "LOOK", "")
    If Val(UserData(ID, 3)) = 0 Then SayFile ID, "misc/helper.txt"
    SayPrompt ID, ""
    Exit Sub
End If

If UserData(ID, 0) = "" Then
    If UCase$(Verb$) = "NEW" Then
        ConnInfo(ID, 1) = "1"
        SayFile ID, "misc/newplayer.txt"
        SayPrompt ID, "Your New Username: "
        Exit Sub
    End If
    
    If UCase$(Verb$) = "GUEST" Then
        ConnInfo(ID, 1) = "4"
        Verb$ = "Guest"
        UserData(ID, 1) = "DUMMY_PASSWORD"
        UserData(ID, 2) = "M"
        UserData(ID, 3) = "0"
        GuestNumber = 1
        For t% = 0 To 64
            If Left$(UCase$(UserData(t%, 0)), Len(Verb$)) = UCase$(Verb$) Then GuestNumber = GuestNumber + 1
        Next t%
        UserData(ID, 0) = Verb$ + Trim$(Str$(GuestNumber))
        If ConnInfo(ID, 2) = "ANSI" Or ConnInfo(ID, 2) = "VT100" Then
            SendData ID, "Your software supports colour."
        Else
             SendData ID, "Your software MAY NOT support colour."
        End If
            
        SayPrompt ID, "Do you want colour (Y/N) ? "
        Exit Sub
    End If

    If PlayerExist(Verb$) = 0 Then
        SayFile ID, "misc\whoareyou.txt"
        SayPrompt ID, "Username: "
        Exit Sub
    End If
    If PlayerExist(Verb$) = 1 Then
        For t% = 0 To 64
            If UCase$(UserData(t%, 0)) = UCase$(Verb$) Then
                SayFile ID, "misc\alreadyin.txt"
                SayPrompt ID, "Username: "
                Exit Sub
            End If
        Next t%
        
        LoadUser ID, Verb$
    End If
    
ElseIf UserTemp(ID) = "PASS" Then
    CheckPass ID, Verb$
Else
    UserData(ID, 12) = Timer
    GetAction ID, UCase$(Verb$), Param$
    SayPrompt ID, ""
End If


End Sub
Sub LoadUser(ID, Username$)

'0 - Player Name
'1 - Player Password
'2 - Player Gender
'3 - Player Points
'4 - Suffix
'5 - Brief
'6 - Profile
'7 - Echo toggle
'8 - VT100 toggle
'9 - Status (0/1/2/3)   0 = Normal   1 = Builder Locs Only   2 = Builder Locs & Items    3 = Immortal
'21 - ANSI Toggle   0 = off    1 = on
'22 - Term width

'-------These Not Used-------
'10 - Location
'11 - Health
'12 - Idler Time
'13 - Login Time
'14 - Opponent
'15 - Strength
'16 - CurrentWeight
'17 - List of Item Flags
'20 - Cached Loc Name
'---------------------------------


Handle = FreeFile
Open App.Path + "\players\" + Username$ + ".plr" For Input Access Read As #Handle
    'Line Input #Handle, UserData(ID, 0)
    'Line Input #Handle, tmp$
    'UserData(ID, 1) = Crypt(tmp$)
    For t% = 0 To 50
        Line Input #Handle, UserData(ID, t%)
    Next t%
Close Handle
UserTemp(ID) = "PASS"
'SendData ID, "Welcome, " + UserData(ID, 0) + "!"
SayPrompt ID, "Password: "
UserData(ID, 12) = Timer

End Sub
Sub SaveUser(ID)
If UserData(ID, 0) = "" Then Exit Sub
Handle = FreeFile
Open App.Path + "\players\" + UserData(ID, 0) + ".plr" For Output Access Write As #Handle
    'Print #Handle, UserData(ID, 0)
    'Print #Handle, UserData(ID, 1)
    For t% = 0 To 50
        Print #Handle, UserData(ID, t%)
    Next t%
Close Handle
End Sub
Sub CheckPass(ID, Password$)
UserTemp(ID) = ""
'Password$ = Crypt(Password$)
Set JC = New JarCrypt
If JC.Decrypt(UserData(ID, 1)) <> Password$ Then
    SendData ID, "Login Incorrect. If you do not have a player account, type NEW to create one."
    SayPrompt ID, "Username: "
    For t% = 0 To 50
        UserData(ID, t%) = ""
    Next t%
    Exit Sub
End If

UserData(ID, 1) = Crypt(Password$)
SendData ID, ""
SendData ID, "Welcome back, " + UserData(ID, 0) + " the " + UserData(ID, 4) + PlayerLevel(ID, Val(UserData(ID, 3)))
If UserData(ID, 3) >= 0 Then SendData ID, "Your current score is " + Format$(Val(UserData(ID, 3)), "###,###,##0")
SendData ID, ""

Call PutPlayerInGame(ID)
Call DoAction(ID, "LOOK", "")

SayPrompt ID, ""
    
End Sub
Sub PutPlayerInGame(ID)

SayFile ID, "misc\motd.txt"

Call ResetPlayerVars(ID, 0)

UserData(ID, 13) = Timer
UserData(ID, 12) = Timer

PlayerLoc = Val(UserData(ID, 10))
GetExits ID, PlayerLoc
SysEmoteToRoom ID, PlayerLoc, "enters the world."
SendDDE UserData(ID, 0) + " enters the world."
Log Trim$(Str$(ID)) + " LOGIN Username: " + UserData(ID, 0)
ConnInfo(ID, 1) = "99"

Call WebWriteWhoInGame
Call DesktopAlert("MUD32", UserData(ID, 0) + " has joined the world. (" + ConnInfo(ID, 0) + ")")
Call UpdateLastOn(UserData(ID, 0))

SayFile ID, "misc\enterworld.txt"

End Sub
Sub ResetPlayerVars(ID, Reason)
'if Reason = 1 then it's a re-incarnate.

UserData(ID, 16) = "0" 'we are no longer carrying anything

If StickyQuit = 1 And Reason = 0 Then
    If UserData(ID, 10) = "" Then UserData(ID, 10) = Trim$(Str$(StartLoc))
Else
    UserData(ID, 10) = Str$(StartLoc)
End If

UserLevel = PlayerLevelNo(UserData(ID, 3)) - 1

'Set Dynamic Health + Strength
UserData(ID, 11) = Levels(UserLevel, 3)  '"100"
UserData(ID, 15) = Levels(UserLevel, 4)  '"1000"

'This is the superuser account so flag it as an immortal
If UCase$(UserData(ID, 0)) = UCase$(SuperUser) Then UserData(ID, 9) = "3"

End Sub
Sub GetAction(ID, UserVerb$, Param$)
UserData(ID, 12) = Timer
Form2.File1.Path = App.Path + "\cmd"
Form2.File1.Pattern = "*.vrb"
Form2.Refresh

x$ = UserData(ID, 10) + ":" + UserData(ID, 0) + "> " + UserVerb$ + " " + Param$
BugPrint x$

If UserVerb$ = "CON" Or UserVerb$ = "PRN" Or Left$(UserVerb$, 3) = "LPT" Or UserVerb$ = "NUL" Then
    SendData ID, "Sorry, that command doesn't work here."
    Exit Sub
End If

'Debug Logging - Shshh, don't tell anyone we log what they type
Handle = FreeFile
Open App.Path + "\logs\" + UserData(ID, 0) + ".log" For Append Access Read Write As Handle
Print #Handle, Date$ + "," + Time$ + "," + UserVerb$ + " " + Param$
Close #Handle

'If UserVerb$ = "HELP" Then
'   For T% = 0 To Form2.File1.ListCount - 1
'       Filename$ = Form2.File1.List(T%)
'        Handle = FreeFile
'        Open App.Path + "\cmd\" + Filename$ For Input Access Read As #Handle
'            Line Input #Handle, ActHelp$
'            Line Input #Handle, Action$
'        Close Handle
'        SendData ID, Left$(UCase$(Left$(Filename$, Len(Filename$) - 4)) + Space$(10), 10) + ActHelp$
'    Next T%

If UserVerb$ = "HELP" Then
    Call HelpSystem(ID, Param$)
Else
    Form2.File1.Pattern = UserVerb$ + ".vrb"
    If Form2.File1.ListCount > 0 Then
        Handle = FreeFile
        Open App.Path + "\cmd\" + UserVerb$ + ".vrb" For Input Access Read As #Handle
            Line Input #Handle, ActHelp$
            Line Input #Handle, Action$
        Close Handle
        
        DoAction ID, Action$, Param$
    Else
        '---- Could Be An Exit Command
        
        PlayerLoc = Val(UserData(ID, 10))
                    
        LocFlags(ID) = ""
        LocFlags(ID) = GetLocFlags(PlayerLoc)
        If InStr(LocFlags(ID), "+DARK") Then LocDark = 1
        If InStr(UserData(ID, 17), "-DARK") Then LocDark = 0
        
        If LocDark = 1 Then MaxExit = 0 Else MaxExit = 9 'Only the first exit will work if it's dark

        For t% = 0 To MaxExit
            If UCase$(PlayerExit(ID, t%, 0)) = UserVerb$ Or UCase$(PlayerExit(ID, t%, 1)) = UserVerb$ Then
                
                If Left$(UserTemp(ID), 6) = "VICTIM" Or Left$(UserTemp(ID), 6) = "ATTACK" Then
                    SendData ID, "You are in the middle of a fight, use FLEE to run away (this will cost you 10% of your score)"
                    Exit Sub
                End If

                'SendData ID, "Valid Exit to " + PlayerExit(ID, T%, 2) + " " + PlayerExit(ID, T%, 3)
                
                If IfExist(Trim$(App.Path + "\locations\" + PlayerExit(ID, t%, 2) + ".loc")) <> True Then
                    SendData ID, "A mysterious force prevents you from going that way."
                    Exit Sub
                End If
                            
                If PlayerExit(ID, t%, 3) = "" Then
                    ExitMess$ = "You go " + PlayerExit(ID, t%, 1) + "..."
                Else
                    ExitMess$ = PlayerExit(ID, t%, 3)
                End If
                SendData ID, ExitMess$
                SendData ID, ""
                If Val(PlayerExit(ID, t%, 2)) <> Val(UserData(ID, 10)) Then
                    If LocDark <> 1 Then
                        SysEmoteToRoom ID, PlayerLoc, "goes " + PlayerExit(ID, t%, 1) + "."
                    End If
                    UserData(ID, 10) = PlayerExit(ID, t%, 2)
                    PlayerLoc = Val(UserData(ID, 10))
                    GetExits ID, PlayerLoc
                    If LocDark <> 0 Then
                        SysEmoteToRoom ID, PlayerLoc, "arrives."
                    End If
                    
                  
                    LookAtLocation ID, PlayerLoc, 0
                
                    WhosHere ID, PlayerLoc
                End If
                Valid = 1
                t% = 10
            End If
        Next t%

        UserData(ID, 12) = Timer
        If Valid <> 1 Then SendData ID, "Sorry, that command doesn't work here."
    End If
End If

End Sub
Function ConvertIP(IPLong)
Reckner = "0123456789ABCDEF"
IPhex = Right$("00000000" + Hex$(IPLong), 8)
For t% = 7 To 0 Step -2
    LumpA$ = Mid$(IPhex, t%, 1)
    LumpB$ = Mid$(IPhex, t% + 1, 1)
    IPValB = InStr(Reckner, LumpB$) - 1
    IPValA = (InStr(Reckner, LumpA$) - 1) * 16
    IPVal$ = IPVal$ + Trim$(Str$(IPValA + IPValB))
    If t% > 2 Then IPVal$ = IPVal$ + "."
Next t%
ConvertIP = IPVal$
End Function

Public Sub SayFile(ID, File$)
x = Dir(App.Path + "\" + File$)
If x = "" Then
    SendData ID, "[Error: File Not Found: " + File$ + "]"
    Exit Sub
End If

Handle = FreeFile
Open App.Path + "\" + File$ For Input Access Read As #Handle
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    If InStr(Ln$, "$") <> 0 Then
        i% = InStr(Ln$, "$OS$")
        If i% <> 0 Then
            'Insert = Environ("OS")
            Ln$ = Left$(Ln$, i% - 1) + WindowsVersion + Mid$(Ln$, i% + 4)
        End If
        i% = InStr(Ln$, "$VERSION$")
        If i% <> 0 Then
            Ln$ = Left$(Ln$, i% - 1) + BuildVersion + Mid$(Ln$, i% + 9)
        End If
    End If
    SendData ID, Ln$
Loop
Close #Handle
End Sub
Sub SayPrompt(ID, Text$)
If Text$ = "" Then Text$ = UserData(ID, 0) + "> "
If ConnInfo(ID, 1) = "99" Then Text$ = AnsiTable(Val(UserData(ID, 21)), 3) + Text$ + AnsiTable(Val(UserData(ID, 21)), 4)
Text$ = Text$ + LineBuffer(ID)

SocketWrite ID, Text$

If UserData(ID, 8) = "1" And ConnInfo(ID, 1) = "99" Then Call ShowVT100Header(ID)
If ConnInfo(ID, 1) = "99" Then
    SocketWrite ID, AnsiTable(Val(UserData(ID, 21)), 4)
End If

End Sub
Sub ResetANSIColor(ID)
If ConnInfo(ID, 1) = "99" Then
    SocketWrite ID, AnsiTable(Val(UserData(ID, 21)), 0)
End If
End Sub
Function PlayerExist(PlayerName$)
    PlayerExist = 0
    x = Dir(App.Path + "\players\" + PlayerName$ + ".plr")
    If x <> "" Then PlayerExist = 1
End Function

Sub ConsoleCommand(ID, Param$)
    OldParam$ = Param$
    Dim PArray(10) As String
    If ID < 1 Then tmp$ = "Console" Else tmp$ = UserData(ID, 0)
    
    Log tmp$ + "> " + Param$
    
    If Param$ = "" Then
        SendData ID, "For help on the ADMIN command type HELP ADMIN."
        Exit Sub
    End If
    Element = 0
    Do
        i% = InStr(Param$, " ")
        If i% <> 0 Then
            PArray(Element) = Trim$(Left$(Param$, InStr(Param$, " ")))
            Param$ = Trim$(Mid$(Param$, InStr(Param$, " ")))
            Element = Element + 1
        Else
            If Param$ <> "" Then PArray(Element) = Param$
        End If
    Loop Until i% = 0 Or Element = 11
       
    'For T% = 0 To 10: Debug.Print PArray(T%): Next T%

    Keyword$ = UCase$(PArray(0))
    
    If Keyword$ = "KICK" Then
        SocketNo = Val(PArray(1))
        SendData SocketNo, "You have been kicked from the server!"
        Call DoAction(SocketNo, "QUIT", "Kicked by " + UserData(ID, 0))
    End If
    
    
    If Keyword$ = "SAYTO" Then
        SendData Val(PArray(1)), "Console Says, " + Mid$(OldParam$, InStr(OldParam$, PArray(1)) + Len(PArray(1)))
        SayPrompt Target, ""
    End If
    
    If Keyword$ = "RESET" Then
        LoadItemArray
        HelpCommands
        If ID > 0 Then Log "Game reset by " + UserData(ID, 0)
        Broadcast "* Due to essentional maintainence, the game has been reset (Sorry)."
    End If
    
    If Keyword$ = "SYSMSG" Then
        Broadcast "* System Message: " + Mid$(OldParam$, InStr(OldParam$, PArray(0)) + Len(PArray(0)))
    End If
    
    If Keyword$ = "SETUSER" Then
        For t% = 0 To 64
            If UCase$(UserData(t%, 0)) = UCase$(PArray(1)) Then
                UserData(t%, 9) = PArray(2)
                Msg$ = SetUserLevel(Val(PArray(2)))
                SendData t%, "You have been granted the status of '" + Msg$ + "'."
                SayPrompt t%, ""
            End If
        Next t%
    End If
    If Keyword$ = "SHUTDOWN" Then
             If ID > 0 Then Log "Remote Server Shutdown by " + UserData(ID, 0)
            ShutdownServer ID, 0
    End If
        
End Sub
Sub ShutdownServer(ID, Interactive)
If Interactive = 1 Then
    x = MsgBox("Are you sure you want to shutdown the server ?", vbQuestion + vbYesNo, App.Title)
    If x = 7 Then Exit Sub
Else
    tmp$ = "non-"
End If


Broadcast "* Server Is Shutting Down!"
If SystemTray = 1 Then
    Call SysTrayIcon_Delete
End If

Log "Server was shutdown"
'Unload Form1

End
End Sub
Function SetUserLevel(LevelNo)
SetUserLevel = "Transient"
If LevelNo = 0 Then SetUserLevel = "Normal User"
If LevelNo = 1 Then SetUserLevel = "Builder"
If LevelNo = 2 Then SetUserLevel = "World Creator"
If LevelNo = 3 Then SetUserLevel = "Immortal"
End Function
Sub xLoadLevelArray()
Log "Loading Level Array..."
Handle = FreeFile
Open App.Path + "\misc\levels.dat" For Input Access Read As Handle
Counter = 0
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    If Trim$(Ln$) <> "" Then
        i% = InStr(Ln$, "#")
        Points = Left$(Ln$, i% - 1)
        Ln$ = Mid$(Ln$, i% + 1)
        
        i% = InStr(Ln$, "#")
        MaleLevel = Left$(Ln$, i% - 1)
        FemaleLevel = Mid$(Ln$, i% + 1)
        
        Levels(Counter, 0) = Points
        Levels(Counter, 1) = MaleLevel
        Levels(Counter, 2) = FemaleLevel
        
        Counter = Counter + 1
    End If
Loop

MaxLevels = Counter - 1
End Sub
Sub LoadLevelArray()
Log "Loading (New) Level Array..."
Dim LevBuff() As String

Handle = FreeFile
Open App.Path + "\misc\levels.dat" For Input Access Read As Handle
Counter = 0
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    If Trim$(Ln$) <> "" Then
        LevBuff() = Split(Ln$, "#")
        
        For t% = 0 To UBound(LevBuff())
            Levels(Counter, t%) = LevBuff(t%)
        Next t%
        
        If Val(Levels(Counter, 3)) < 1 Then Levels(Counter, 3) = "100"
        If Val(Levels(Counter, 4)) < 1 Then Levels(Counter, 4) = "1000"
        
        Counter = Counter + 1
    End If
Loop

MaxLevels = Counter - 1
End Sub
Function PlayerLevel(ID, Points)
If ID = -1 Then Gender$ = "M"
If ID = -2 Then Gender$ = "F"
If ID >= 0 Then Gender$ = UserData(ID, 2)

'If ID >= 0 Then
'    If Val(UserData(ID, 9)) = 3 Then
'        PlayerLevel = "Admin"
'        Exit Function
'    End If
'End If

'If Points = -999999 Then
'    PlayerLevel = GodLevel  ' bit of a fudge - come back and fix it.
'    Exit Function
'End If

If Points < 0 Then
    PlayerLevel = "Vagrant"
    Exit Function
End If

For t% = 0 To 20
    If Points < Val(Levels(t% + 1, 0)) Then
        If UCase$(Gender$) = "M" Then
            PlayerLevel = Levels(t%, 1)
        Else
            PlayerLevel = Levels(t%, 2)
        End If
        Exit Function
    End If
    If Levels(t%, 1) = "" Then
        If UCase$(Gender$) = "M" Then
            PlayerLevel = Levels(t% - 1, 1)
        Else
            PlayerLevel = Levels(t% - 1, 2)
        End If
        Exit Function
    End If
Next t%
PlayerLevel = "Player"
End Function
Function PlayerLevelNo(Points)
PlayerLevelNo = 1

For t% = 1 To 20
    If Points < Val(Levels(t% + 1, 0)) Then
        PlayerLevelNo = t%
        Exit Function
    End If
    If Levels(t%, 1) = "" Then
        PlayerLevelNo = Levels(t% - 1, 1)
        Exit Function
    End If
Next t%
End Function
Function CheckLocZone(PlayerLoc)
CheckLocZone = "NULL"
Handle = FreeFile
x = Dir("locations\" + Trim$(Str$(PlayerLoc)) + ".loc")

If x = "" Then
    CheckLocZone = "_BADEXIT"
    Exit Function
End If

Open "locations\" + Trim$(Str$(PlayerLoc)) + ".loc" For Input Access Read As Handle
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    LnType$ = Left$(Ln$, 1)
    If LnType$ = "4" Then
        CheckLocZone = UCase$(Mid$(Ln$, 3))
        Exit Do
    End If
Loop

Close Handle
End Function
Sub GetExits(ID, PlayerLoc)
For t% = 0 To 9
    For r% = 0 To 3
        PlayerExit(ID, t%, r%) = ""
    Next r%
Next t%
   
Handle = FreeFile
Open "locations\" + Trim$(Str$(PlayerLoc)) + ".loc" For Input Access Read As Handle
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    LnType$ = Left$(Ln$, 1)
    Ln$ = Mid$(Ln$, 3)
    If LnType$ = "2" And Counter < 10 Then
        i% = InStr(Ln$, ",")
        PlayerExit(ID, Counter, 0) = Trim$(Left$(Ln$, i% - 1))
        Ln$ = Mid$(Ln$, i% + 1)
        i% = InStr(Ln$, ",")
        PlayerExit(ID, Counter, 1) = Trim$(Left$(Ln$, i% - 1))
        Ln$ = Mid$(Ln$, i% + 1)
        i% = InStr(Ln$, ",")
        If i% <> 0 Then
            PlayerExit(ID, Counter, 2) = Trim$(Left$(Ln$, i% - 1))
            PlayerExit(ID, Counter, 3) = UCase$(Left$(Trim$(Mid$(Ln$, i% + 1)), 1)) + Mid$(Trim$(Mid$(Ln$, i% + 1)), 2)
        Else
            PlayerExit(ID, Counter, 2) = Trim$(Ln$)
        End If
        Counter = Counter + 1
    End If
Loop

Close Handle
TotalPlayerExit(ID) = Counter
End Sub
Function IfExist(Filename$)
On Error GoTo IfExistTrap
fail = 0
Handle = FreeFile
Open Filename$ For Input Access Read As Handle
Close Handle
If fail = 0 Then IfExist = True
On Error GoTo 0
Exit Function

IfExistTrap:
fail = 1: Resume Next

End Function

Function Crypt(Pass$)
'For t% = 1 To Len(Pass$)
'    x% = Asc(Mid$(Pass$, t%, 1))
'    y% = x% Xor (27 + t%)
'    NewPass$ = NewPass$ + Chr$(y%)
'Next t%

Set uJarCrypt = New JarCrypt

Crypt = uJarCrypt.Encrypt(Pass$, 32)

End Function



Sub SetBit(Bit, BitString, Toggle)
Offset = (Int(Bit / 8)) + 1
BitNo = Bit - (Int(Bit / 8) * 8)
Char = Mid$(BitString, Offset, 1)
CharVal = Asc(Char)
BitVal = 2 ^ BitNo
If CharVal And BitVal Then
    If Toggle = 0 Then
        CharVal = CharVal - BitVal
    End If
Else
    If Toggle = 1 Then
        CharVal = CharVal + BitVal
    End If
End If
Mid$(BitString, Offset, 1) = Chr$(CharVal)

End Sub
Function GetBit(Bit, BitString)
Offset = (Int(Bit / 8)) + 1
BitNo = Bit - (Int(Bit / 8) * 8)
Char = Mid$(BitString, Offset, 1)
CharVal = Asc(Char)
BitVal = 2 ^ BitNo
If CharVal And BitVal Then
    GetBit = True
Else
    GetBit = False
End If

End Function


Sub LoadItemArray()
Log "Flushing Item Array..."
For t% = 0 To 1000
    ItemArray(t%, 0) = "0"
Next t%
SuperValue = 0
Log "Restacking Item Array..."
Form2.File1.Path = App.Path + "\items"
Form2.File1.Pattern = "*.itx"
For t% = 0 To Form2.File1.ListCount - 1
    ItemFile$ = Trim$(Form2.File1.List(t%))
    ItemVal% = Val(Left$(ItemFile$, Len(ItemFile$) - 4))
    ItemFile$ = "items\" + ItemFile$
      
    ItemArray(ItemVal%, 0) = "1" ' State: 1=in world 2=in hand 0=destroyed
    ItemArray(ItemVal%, 1) = ReadINI(ItemFile$, "location") ' Location
    ItemArray(ItemVal%, 2) = ReadINI(ItemFile$, "shortname") ' Short Name
    ItemArray(ItemVal%, 3) = ReadINI(ItemFile$, "fullname") ' Long Name
    ItemArray(ItemVal%, 4) = ReadINI(ItemFile$, "weight") ' Weight
    ItemArray(ItemVal%, 5) = ReadINI(ItemFile$, "value") ' Value
    ItemArray(ItemVal%, 6) = ReadINI(ItemFile$, "damage") ' Damage
    ItemArray(ItemVal%, 7) = ReadINI(ItemFile$, "flags") ' flags: -dark -poison -water etc
    
    SuperValue = SuperValue + Val(ItemArray(ItemVal%, 5))
    
    'LoadItemData 0, Item$
    'ItemArray(ItemVal%, 0) = "1"
    'ItemArray(ItemVal%, 1) = ItemData(0, 7) ' Location
    'ItemArray(ItemVal%, 2) = ItemData(0, 3) ' Short Name
    'ItemArray(ItemVal%, 3) = ItemData(0, 0) ' Long Name
    'ItemArray(ItemVal%, 4) = ItemData(0, 4) ' Weight
    'ItemArray(ItemVal%, 5) = ItemData(0, 5) ' Value
    'ItemArray(ItemVal%, 6) = ItemData(0, 6) ' Damage
    
Next t%

'Reset weight carried for any players that are logged in.
For t% = 0 To 64
    If ConnInfo(t%, 1) <> "" Then UserData(t%, 16) = 0
Next t%

Log "Total Value Of Items is" + Str$(SuperValue)

End Sub
Sub LoadItemData(ID, Item$)
    Handle = FreeFile
    Open App.Path + "\items\" + Item$ For Input Access Read As Handle
    Do Until EOF(Handle)
        Line Input #Handle, Ln$
        Ln$ = Trim$(Ln$)
        If Ln$ <> "" Then
            LineTag = Val(Left$(Ln$, 1))
            Ln$ = Mid$(Ln$, 3)
            ItemData(ID, LineTag) = Ln$
        End If
    Loop
    Close Handle
End Sub




Function ReadINI(Filename$, Keyword$)

' Called by: Value = ReadINI(INIFile, Keyword)
' this is not good for massive INI files, but will suffice.
' Yes I know about ReadPrivateProfile etc, but thats not very portable is it?

Handle = FreeFile
Open App.Path + "\" + Filename$ For Input Access Read As #Handle
Do Until EOF(Handle)
Line Input #Handle, Ln$
Ln$ = Trim$(Ln$)
If Ln$ <> "" And Left$(Ln$, 1) <> ";" And Left$(Ln$, 1) <> "[" Then
    If UCase$(Left$(Ln$, Len(Keyword$))) = UCase$(Keyword$) Then
        i% = InStr(Ln$, "=")
        tmp$ = Mid$(Ln$, i% + 1)
        ReadINI = Trim$(tmp$)
        Close Handle
        Exit Function
    End If
End If
Loop
ReadINI = ""
Close Handle

End Function

Sub HelpSystem(ID, Param$)
HelpFileSuffix$ = ".txt"
HelpFile$ = "help\main"
If Param$ <> "" Then HelpFile$ = "help\" + Trim$(Param$)

If UCase$(Param$) = "IMMORTAL" And Val(UserData(ID, 9)) <> 3 Then
    HelpFile$ = "help\MORTAL"
End If

If IfExist(App.Path + "\" + HelpFile$ + HelpFileSuffix$) = True Then
    SayFile ID, HelpFile$ + HelpFileSuffix$
Else
    Stat$ = "Sorry, no help available on that topic."
    
    If Param$ <> "" Then HelpFile$ = "cmd\" + Trim$(Param$) + ".vrb"
    If IfExist(App.Path + "\" + HelpFile$) = True Then
        Handle = FreeFile
        Open HelpFile$ For Input Access Read As #Handle
        Line Input #Handle, Ln$
        Stat$ = Ln$
        Close Handle
        Stat$ = Replace(Stat$, "\n", vbCrLf)
    End If
    
    If UCase$(Param$) = "VERSION" Then
        SendData ID, "MUD32 Engine Version " + BuildVersion + " on " + WindowsVersion + "."
        Stat$ = ReadINI("mud32.ini", "ADMINMSG")
    End If
    
    SendData ID, Stat$
    
End If
End Sub
Sub HelpCommands()
Offset = 0
CmdList$ = ""
Form2.File1.Path = App.Path + "\cmd"
Form2.File1.Pattern = "*.vrb"
Form2.Refresh
For t% = 0 To Form2.File1.ListCount - 1
    CmdList$ = CmdList$ + Left$(LCase$(Left$(Form2.File1.List(t%), Len(Form2.File1.List(t%)) - 4)) + Space$(12), 12)
    Offset = Offset + 1
    If Offset = 6 Then Offset = 0: CmdList$ = Trim$(CmdList$) + vbCrLf
Next t%
Handle = FreeFile
Open App.Path + "\help\commands.txt" For Output Access Write As #Handle
Print #Handle, CmdList$
CmdList$ = CmdList$ + vbTab
Close Handle

End Sub


Sub RespawnItems()
    Dim Newloc() As String
    For t% = 0 To 1000
        If RespawnList(t%, 0) > -1 Then
            RespawnList(t%, 1) = RespawnList(t%, 1) - 1
            If RespawnList(t%, 1) <= 0 Then
                'respawn it
                ItemFile$ = "items\" + Trim$(Str$(RespawnList(t%, 0))) + ".itx"
                OrigLoc$ = ReadINI(ItemFile$, "location")
                Newloc() = Split(OrigLoc$, ",")
                uloc% = Int(Rnd(1) * (UBound(Newloc()) + 1))
                Log "RESPAWN: Respawning '" + ItemArray(RespawnList(t%, 0), 3) + "' at location " + Newloc(uloc%)
                ItemArray(RespawnList(t%, 0), 0) = "1"
                ItemArray(RespawnList(t%, 0), 1) = Newloc(uloc%) 'OrigLoc$
                Mess$ = ReadINI(ItemFile$, "LocDesc")
                SysMsgToRoom Val(OrigLoc$), Mess$
                RespawnList(t%, 0) = -1
            End If
        Else
            If RespawnList(t%, 0) = 0 Then t% = 1001
        End If
    Next t%
    
End Sub

Sub AddToRList(ItemNo)
    ItemFile$ = "items\" + Trim$(Str$(ItemNo)) + ".itx"
    For t% = 0 To 1000
        If RespawnList(t%, 0) = -1 Or RespawnList(t%, 0) = 0 Then
            RespawnList(t%, 0) = ItemNo
            ForceMin = Val(ReadINI(ItemFile$, "respawn"))
            If ForceMin < MinRespawn Then ForceMin = MinRespawn
            RandomSpawn = (Int(Rnd(1) * (MaxRespawn - ForceMin)) + ForceMin)
            RespawnList(t%, 1) = RandomSpawn
            Log "RESPAWN: '" + ItemArray(ItemNo, 3) + "' in" + Str$(RandomSpawn) + " mins."
            t% = 1001
        End If
    Next t%
            
End Sub

Sub SendDDE(DDEMessage)
If DDEReporting <> 1 Then Exit Sub
On Error GoTo DDEFail

Form2.Text1.Text = DDEMsgPrefix + " " + DDEMessage
Form2.Text1.LinkTopic = DDETopic
If DDEItem <> "NULL" Then Form2.Text1.LinkItem = DDEItem
Form2.Text1.LinkMode = 2
Form2.Text1.LinkPoke
Form2.Text1.LinkMode = 0

DDEFail:
On Error GoTo 0
End Sub

Sub CreateTopTenFile()
For t% = 0 To 10
    TopTenArray(t%, 1) = Str$(0)
Next t%

Handle = FreeFile
Open App.Path + "\misc\top10.lst" For Output Access Write As #Handle
For t% = 1 To 10
    Ln$ = TopTenArray(t%, 0) + "," + TopTenArray(t%, 1) + "," + TopTenArray(t%, 2) + "," + TopTenArray(t%, 3)
    Print #Handle, Ln$
Next t%
Close #Handle


End Sub
Sub UpdateTopTen(Name$, Score$, Gender$)
If UCase$(Left$(Name$, 5)) = "GUEST" Then Exit Sub

UKDate$ = Mid$(Date$, 4, 3) + Mid$(Date$, 1, 3) + Right$(Date$, 4)

    Handle = FreeFile
   Open App.Path + "\misc\top10.lst" For Input Access Read As #Handle
    Count = 1
    Do Until EOF(Handle)
        Input #Handle, AName$, AScore$, AGender$, ALastPlay$
       TopTenArray(Count, 0) = AName$
        TopTenArray(Count, 1) = AScore$
        TopTenArray(Count, 2) = AGender$
        TopTenArray(Count, 3) = ALastPlay$
        Count = Count + 1
        If Count = 11 Then Exit Do
    Loop
    Close #Handle
    
    For x% = 1 To 10
        If TopTenArray(x%, 0) = Name$ Then
            For r% = x% To 9
                For C% = 0 To 3
                    TopTenArray(r%, C%) = TopTenArray(r% + 1, C%)
                Next C%
            Next r%
            TopTenArray(10, 0) = ""
            TopTenArray(10, 1) = ""
            TopTenArray(10, 2) = "0"
            TopTenArray(10, 3) = ""
            
            x% = 11
        End If
    Next x%
    
    For r% = 10 To 1 Step -1
        If Val(Score$) > Val(TopTenArray(r%, 1)) Then
            If Val(Score$) <= Val(TopTenArray(r% - 1, 1)) Or r% = 1 Then
                For s% = 9 To r% Step -1
                    For C% = 0 To 3
                        TopTenArray(s% + 1, C%) = TopTenArray(s%, C%)
                    Next C%
                Next s%
                TopTenArray(r%, 0) = Name$
                TopTenArray(r%, 1) = Score$
                TopTenArray(r%, 2) = Gender$
                TopTenArray(r%, 3) = Format$(UKDate$, "dd mmm yyyy")
            End If
        End If
    Next r%
    
    Handle = FreeFile
    Open App.Path + "\misc\top10.lst" For Output Access Write As #Handle
    For t% = 1 To 10
        Ln$ = TopTenArray(t%, 0) + "," + TopTenArray(t%, 1) + "," + TopTenArray(t%, 2) + "," + TopTenArray(t%, 3)
        Print #Handle, Ln$
    Next t%
    Close #Handle
    Call WebWriteTopTen
End Sub
Sub WebWriteTopTen()
UKDate$ = Mid$(Date$, 4, 3) + Mid$(Date$, 1, 3) + Right$(Date$, 4)

WHandle = FreeFile
Open WebRoot + "\TopTen.htm" For Output Access Write As WHandle
Print #WHandle, "<html><body><pre>"
Print #WHandle, "Top Ten Players on " + Format$(UKDate$, "dd mmm yyyy")
Print #WHandle, ""
Print #WHandle, "Pos  Name                    Score  Level             Last Played"
Print #WHandle, "-----------------------------------------------------------------"
Handle = FreeFile
Open App.Path + "\misc\top10.lst" For Input Access Read As #Handle
For t% = 1 To 10
    Input #Handle, Name$, Score$, Gender$, Played$
    If Gender$ = "M" Then Offset = -1 Else Offset = -2
    Level$ = PlayerLevel(Offset, Val(Score$))
    Ln$ = Right$("  " + Trim$(Str$(t%)), 2) + "   " + Left$(Name$ + Space$(20), 17) + Right$(Space$(12) + Format$(Val(Score$), "###,###,##0"), 12) + "  " + Left$(Level$ + Space$(20), 18) + Played$
    If Name$ <> "" Then Print #WHandle, Ln$
Next t%
Close #Handle
Print #WHandle, "-----------------------------------------------------------------"
Print #WHandle, ""
Print #WHandle, "</pre></body></html>"
Close #WHandle
End Sub
Sub WebWriteWhoInGame()
    UKDate$ = Mid$(Date$, 4, 3) + Mid$(Date$, 1, 3) + Right$(Date$, 4)

    WHandle = FreeFile
    Open WebRoot + "\Who.htm" For Output Access Write As WHandle
    Print #WHandle, "<html><body><pre>"
    Print #WHandle, "Current Players on " + Format$(UKDate$, "dd mmm yyyy") + " at " + Time$
        
    Header$ = Left$("Player Name" + Space$(50), 50) + "Status / Online Time"
    Print #WHandle, Header$
    Header$ = Left$("-----------" + Space$(50), 50) + "--------------------"
    Print #WHandle, Header$
    
    ActivePorts = 0
    For t% = 0 To UBound(LineBuffer, 1)
        If UserData(t%, 0) <> "" Then
            Idle$ = ""
            IdleTime = (Timer - Val(UserData(t%, 12))) / 60
            If IdleTime < 0 Then IdleTime = 0
            ConnTime = (Timer - Val(UserData(t%, 13))) / 60
            FullName$ = UserData(t%, 0) + " the " + UserData(t%, 4) + PlayerLevel(t%, Val(UserData(t%, 3)))
        
            If Int(IdleTime) <> 0 Then
                Idle$ = "Idle " + Trim$(Str$(Int(IdleTime))) + " mins"
            Else
                Idle$ = "Active"
            End If
            Conn$ = Trim$(Str$(Int(ConnTime))) + " mins"
            ActivePorts = ActivePorts + 1
            PADName$ = Left$(FullName$ + Space$(50), 50)
            Txt$ = PADName$ + Idle$ + " / " + Conn$
            Print #WHandle, Txt$
        End If
    Next t%
    Footer$ = Trim$(Str$(ActivePorts)) + " Connected Players."
    Print #WHandle, String$(Len(Footer$), "-")
    Print #WHandle, Trim$(Str$(ActivePorts)) + " Connected Players."
    Close #WHandle
End Sub



Function Success(ID, Handicap, CurrPoints)
    
    'With a Handicap of 50,
    'and AttackerLevel of 1,  will return 50/50
    'or AttackerLevel of MaxLevels will return 100%
    'this means you can have an adjusting success/fail ratio based on level.
    
    TotalPlayerLevels = MaxLevels
    AttackerLevel = PlayerLevelNo(CurrPoints)

    Perc = (100 / TotalPlayerLevels)
    Ae = CInt(Perc * AttackerLevel) + Handicap
    HitBar = Left$(String$(Ae, "H") + String$(100, "M"), 100)
    i% = Int(Rnd(1) * 99) + 1
    Checkit = Mid$(HitBar, i%, 1)
    If Checkit = "H" Then
        Success = True
    Else
        Success = False
    End If

End Function
Sub SetupVT100Codes()
ESC$ = Chr$(27)

VTBold(0, 1) = ""
VTBold(0, 0) = ""

VTBold(1, 1) = ESC$ + "[1m"  ' Bold On
VTBold(1, 0) = ESC$ + "[0m" ' Bold Off

End Sub

Sub ShowVT100Header(ID)
'Oh how evil is this! - Jar, 1st July 2003
ESC$ = Chr$(27)
SaveCur$ = ESC$ + "7"
RestCur$ = ESC$ + "8"
GotoTopLeft$ = ESC$ + "[1;1H"
If UserData(ID, 21) = 0 Then
    InvOn$ = ESC$ + "[7m"
    InvOff$ = ESC$ + "[0m"
Else
    InvOn$ = ESC$ + "[0;47;34m"
    InvOff$ = ESC$ + "[40;37m"
End If

Header$ = Left$(UserData(ID, 20) + Space$(35), 35)
Header$ = Header$ + Left$("Health: " + Trim$(UserData(ID, 11)) + "  Strength: " + Trim$(Str$(UserData(ID, 15))) + Space$(40), 30)
Header$ = Header$ + Right$(Space$(15) + "Score: " + Format$(Val(UserData(ID, 3)), "###,###,##0"), 15)

Mess$ = SaveCur$ + GotoTopLeft$ + InvOn$ + Header$ + InvOff$ + RestCur$

SocketWrite ID, Mess$

End Sub


Sub RebuildUsedLocCache()
Handle = FreeFile
Open App.Path + "\locations\usedcache.dat" For Binary Access Write As Handle
'For T% = 1 To 9999
    Put #Handle, 1, String$(9999, "0")
    'DoEvents
'Next T%
Close #Handle

Form2.File1.Path = App.Path + "\locations"
Form2.File1.Pattern = "*.loc"
For t% = 0 To Form2.File1.ListCount - 1
    FName$ = Left$(Form2.File1.List(t%), Len(Form2.File1.List(t%)) - 4)
    FNameVal = Val(FName$)
    Call ToggleLocCache(FNameVal, "1")
Next t%
Log "Cache rebuild complete, " + Trim(Str(Form2.File1.ListCount)) + " locations verified."

End Sub

Sub ToggleLocCache(RecNo, RecChar$)
Handle = FreeFile
Open App.Path + "\locations\usedcache.dat" For Binary Access Write As Handle
Put #Handle, RecNo, RecChar$
Close #Handle

End Sub
Function GetLocCache(RecNo)
Handle = FreeFile
RecChar$ = " "
Open App.Path + "\locations\usedcache.dat" For Binary Access Read As Handle
Get #Handle, RecNo, RecChar$
GetLocCache = RecChar$
Close #Handle
End Function

Sub WriteLoc(LocNo, Entry$, Param$)
Dim LocTemp(20) As String
LocFile$ = Trim$(Str$(LocNo)) + ".loc"
Handle = FreeFile

' read it in
Open App.Path + "\locations\" + LocFile$ For Input Access Read As #Handle
Counter = 0
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    LocTemp(Counter) = Ln$
    Counter = Counter + 1
Loop
Close #Handle

'update entry
Matched = 0
For t% = 0 To 20
    If Left$(LocTemp(t%), 1) = Entry$ Then
        Matched = 1
        LocTemp(t%) = Entry$ + " " + Param$
        t% = 21
    End If
Next t%
If Matched = 0 Then LocTemp(Counter) = Entry$ + " " + Param$

'write it out
Open App.Path + "\locations\" + LocFile$ For Output Access Write As #Handle
For t% = 0 To 20
    If LocTemp(t%) <> "" Then
        Print #Handle, LocTemp(t%)
    End If
Next t%
Close #Handle

End Sub


Sub CreateLoc(LocNo, LocName$, Owner$)
LocFile$ = Trim$(Str$(LocNo)) + ".loc"
Handle = FreeFile

'write it out
Open App.Path + "\locations\" + LocFile$ For Output Access Write As #Handle
Print #Handle, "0 " + LocName$
Print #Handle, "3 " + Owner$
Close #Handle

End Sub

Function GetLocEntry(LocNo, Entry$)
LocFile$ = Trim$(Str$(LocNo)) + ".loc"
Handle = FreeFile
GetLocEntry = ""

' read it in - and find entry
Open App.Path + "\locations\" + LocFile$ For Input Access Read As #Handle

Do Until EOF(Handle)
    Line Input #Handle, Ln$
    If UCase$(Left$(Ln$, Len(Entry$))) = UCase$(Entry$) Then GetLocEntry = Trim$(Mid$(Ln$, Len(Entry$) + 1))
Loop
Close #Handle

End Function
Sub WriteLocExit(LocNo, Param$)
Dim LocTemp(20) As String
LocFile$ = Trim$(Str$(LocNo)) + ".loc"
Handle = FreeFile

' read it in
Open App.Path + "\locations\" + LocFile$ For Input Access Read As #Handle
Counter = 0
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    LocTemp(Counter) = Ln$
    Counter = Counter + 1
Loop
Close #Handle

'update entry
Matched = 0
For t% = 0 To 20
    If Left$(LocTemp(t%), 1) = "2" Then
        ExitLine$ = Trim$(Mid$(LocTemp(t%), 3))
        i% = InStr(ExitLine$, ",")
        ExitLine$ = Left$(ExitLine$, i% - 1)
        If UCase$(ExitLine$) = UCase$(Left$(Param$, Len(ExitLine$))) Then
            Matched = 1
            LocTemp(t%) = "2 " + Param$
            t% = 21
        End If
    End If
Next t%
If Matched = 0 Then LocTemp(Counter) = "2 " + Param$

'write it out
Open App.Path + "\locations\" + LocFile$ For Output Access Write As #Handle
For t% = 0 To 20
    If LocTemp(t%) <> "" Then
        Print #Handle, LocTemp(t%)
    End If
Next t%
Close #Handle

End Sub

Sub RemoveLocExit(ID, LocNo, Param$)
Dim LocTemp(20) As String
LocFile$ = Trim$(Str$(LocNo)) + ".loc"
Handle = FreeFile

' read it in
Open App.Path + "\locations\" + LocFile$ For Input Access Read As #Handle
Counter = 0
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    LocTemp(Counter) = Ln$
    Counter = Counter + 1
Loop
Close #Handle

'update entry
Matched = 0
For t% = 0 To 20
    If Left$(LocTemp(t%), 1) = "2" Then
        ExitLine$ = Trim$(Mid$(LocTemp(t%), 3))
        i% = InStr(ExitLine$, ",")
        ExitLine$ = Left$(ExitLine$, i% - 1)
        If UCase$(ExitLine$) = UCase$(Left$(Param$, Len(ExitLine$))) Then
            Matched = 1
            LocTemp(t%) = ""
            t% = 21
        End If
    End If
Next t%
If Matched = 0 Then
    SendData ID, "Exit not found in location" + Str$(LocNo) + "."
    Exit Sub
Else
    SendData ID, "Exit removed from location" + Str$(LocNo) + "."
    
End If

'write it out
Open App.Path + "\locations\" + LocFile$ For Output Access Write As #Handle
For t% = 0 To 20
    If LocTemp(t%) <> "" Then
        Print #Handle, LocTemp(t%)
    End If
Next t%
Close #Handle
PlayerLoc = Val(UserData(ID, 10))
If LocNo = PlayerLoc Then GetExits ID, PlayerLoc
End Sub

Sub SetupAnsiCodes(TFile$)

AnsiTable(1, 0) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "Default_Text")))
AnsiTable(1, 1) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "Location_Name")))
AnsiTable(1, 2) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "Location_Description")))
AnsiTable(1, 3) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "Player_Prompt")))
AnsiTable(1, 4) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "Player_Text")))
AnsiTable(1, 5) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "System_Message")))
AnsiTable(1, 6) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "Say_Text")))
AnsiTable(1, 7) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "Emote_Text")))
AnsiTable(1, 8) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "Item_Description")))
AnsiTable(1, 9) = DeESC(ReadINI(TFile$, ReadINI(TFile$, "NPC_Text")))

'For t% = 0 To 15
    'CodeLookup$ = Right$("0" + Trim$(Str$(t%)), 2)
    'AnsiCodes(t%) = DeESC(ReadINI(TFile$, CodeLookup$))
'Next t%
Counter = 0
Handle = FreeFile
Open TFile$ For Input Access Read As #Handle
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    Ln$ = Trim$(Ln$)
    If Reading = 1 And Ln$ <> "" Then
        i% = InStr(Ln$, "=")
        If Left$(Ln$, 1) <> ";" And i% <> 0 Then
            ColID$ = Left$(Ln$, i% - 1)
            ColCode$ = Mid$(Ln$, i% + 1)
            AnsiCodes(Counter, 0) = "&" + ColID$
            AnsiCodes(Counter, 1) = DeESC(ColCode$)
            Counter = Counter + 1
        End If
    End If
    
    
    If Left$(Ln$, 1) = "[" Then
        If UCase$(Ln$) = "[ANSICODES]" Then
            Reading = 1
        Else
            Reading = 0
        End If
        
    End If

Loop
Close Handle


End Sub

Function DeESC(MyStr$)
    MyStr$ = " " + MyStr$ + " "
    Do
        i% = InStr(MyStr$, "{ESC}")
        If i% <> 0 Then
            MyStr$ = Left$(MyStr$, i% - 1) + Chr$(27) + Mid$(MyStr$, i% + 5)
        End If
    Loop Until i% = 0
    DeESC = Trim$(MyStr$)
End Function

Sub InsertColor(ID, TextLine, StripCount)

TextLine = " " + TextLine + " "
StripCount = 0
Do
    i% = InStr(TextLine, "&")
    If i% = 0 Then Exit Do
    PartA$ = Left$(TextLine, i% - 1)
    PartB$ = Mid$(TextLine, i% + 3)
    ColorCode$ = Mid$(TextLine, i%, 3)
    If Val(UserData(ID, 21)) <> "0" Then 'if colour is enabled for this player
        MyCode$ = ""
        For t% = 0 To 15
            'CodeLookup$ = "&" + Right$("0" + Trim$(Str$(t%)), 2)
            
            
            If ColorCode$ = AnsiCodes(t%, 0) Then
                MyCode$ = AnsiCodes(t%, 1)
                StripCount = StripCount + Len(MyCode$)
                t% = 17
            End If
        Next t%
       
    Else
        MyCode$ = ""
    End If
    TextLine = PartA$ + MyCode$ + PartB$
Loop
TextLine = Mid$(TextLine, 2, Len(TextLine) - 2)

End Sub

Sub UpdateConnCount(Index)
Form1.Label3.Caption = Form1.Label3.Caption + Index
If Index > 0 Then Form1.Label4.Caption = Form1.Label4.Caption + Index
Call SystrayIcon_SetTip(Form1.Caption + " - " + Form1.Label3.Caption + Form1.Label5.Caption)
End Sub

Function ChangePlayerPass(PlrName, uOldPass$, uNewPass$, PreCheck)
Dim FileBuff(50) As String
OldPass = uOldPass$
NewPass = Crypt(uNewPass$)
ChangePlayerPass = ""
PlayerFile$ = App.Path + "\players\" + PlrName + ".plr"

If Not IfExist(PlayerFile$) Then
    ChangePlayerPass = "Player does not exist"
    Exit Function
End If

Handle = FreeFile
Counter = 0
Open PlayerFile$ For Input Access Read As #Handle
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    FileBuff(Counter) = Ln$
    Counter = Counter + 1
Loop
Close #Handle
Set JC = New JarCrypt
If JC.Decrypt(FileBuff(1)) <> OldPass And PreCheck = 1 Then
    ChangePlayerPass = "Old Password wrong. Please try again."
    Exit Function
End If
FileBuff(1) = NewPass
Handle = FreeFile
Open PlayerFile$ For Output Access Write As #Handle
For t% = 0 To 50
    Print #Handle, FileBuff(t%)
Next t%
Close #Handle
ChangePlayerPass = NewPass

End Function

Function GetUptime() As String
    Dim ElapsedTime As Double

    ElapsedTime = Abs(Now - MudStartTime)

    UpTimeTmp = ""

    If Int(ElapsedTime) > 0 Then
        UpTimeTmp = UpTimeTmp & Int(ElapsedTime) & " days, "
    End If

    UpTimeTmp = UpTimeTmp & Format$(ElapsedTime, "h") & " hours, "
    UpTimeTmp = UpTimeTmp & Format$(ElapsedTime, "N") & " minutes, "
    UpTimeTmp = UpTimeTmp & Format$(ElapsedTime, "S") & " seconds. "

    GetUptime = UpTimeTmp
End Function
Function GetUptimeConsole() As String
    Dim ElapsedTime As Double

    ElapsedTime = Abs(Now - MudStartTime)

    UpTimeTmp = ""

    If Int(ElapsedTime) > 0 Then
        UpTimeTmp = UpTimeTmp & Int(ElapsedTime) & "d "
    End If

    UpTimeTmp = UpTimeTmp & Format$(ElapsedTime, "h") & "h "
    UpTimeTmp = UpTimeTmp & Format$(ElapsedTime, "N") & "m "
    UpTimeTmp = UpTimeTmp & Format$(ElapsedTime, "S") & "s"

    GetUptimeConsole = UpTimeTmp
End Function

Function GetLocFlags(Location)
GetLocFlags = ""
Handle = FreeFile
Open App.Path + "\locations\" + Trim$(Str$(Location)) + ".loc" For Input Access Read As Handle
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    LnType$ = Left$(Ln$, 1)
    If LnType$ = 5 Then
        GetLocFlags = UCase$(Ln$)
        Exit Do
    End If
Loop
Close Handle

End Function

Function IsBanned(IncomingIP)
IsBanned = False
x = Dir(App.Path + "\misc\banned.dat")
If x = "" Then Exit Function
Handle = FreeFile
Open App.Path + "\misc\banned.dat" For Input Access Read As Handle
Do Until EOF(Handle)
    Line Input #Handle, TestIP
    TestIP = Trim$(TestIP)
    If Left$(TestIP, 1) <> "#" And TestIP <> "" Then
        If Left$(IncomingIP, Len(TestIP)) = TestIP Then
            IsBanned = True
            Exit Do
        End If
    End If
Loop
Close Handle

End Function

Function InvalidName(Ln$)
InvalidName = False
uName$ = Trim$(UCase$(Ln$))
If Len(uName$) < 4 Or Len(uName$) > 20 Then
    InvalidName = True
    Exit Function
End If

Handle = FreeFile
Open App.Path + "\misc\badnames.dat" For Binary Access Read As Handle
uBuffer$ = Space$(LOF(Handle))
Get #Handle, , uBuffer$
BadNames$ = UCase$(uBuffer$)
Close Handle
i% = InStr(BadNames$, uName$)

If i% <> 0 Then InvalidName = True
    
End Function

