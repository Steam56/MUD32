Attribute VB_Name = "Module2"
Sub DoAction(ID, Action$, Param$)



If Action$ = "QUIT" Or Action$ = "LOGOFF" Or Action$ = "LOGOUT" Then
    If Left$(UserTemp(ID), 6) = "VICTIM" Or Left$(UserTemp(ID), 6) = "ATTACK" Then
        Call Flee(ID)
    End If
    DropItem ID, "ALL"
    SayFile ID, "misc\goodbye.txt"
    PlayerLoc = Val(UserData(ID, 10))
    If ConnInfo(ID, 1) = "99" Then
        If Param$ = "" Or Param$ = "Break!" Then
            Mess$ = "leaves the world."
            If Param$ = "" Then Param$ = "Requested."
        Else
            Mess$ = "says '" + Param$ + "', and leaves the world."
        End If
        SysEmoteToRoom ID, PlayerLoc, Mess$
        SendDDE UserData(ID, 0) + " leaves the world."
        Call DesktopAlert("MUD32", UserData(ID, 0) + " has left the world. (" + ConnInfo(ID, 0) + ")")
    End If
    Log Trim$(Str$(ID)) + " QUIT Request: " + UserData(ID, 0) + " (" + Param$ + ")"
    Call SocketClose(ID)
    Call UpdateConnCount(-1)
    ConnInfo(ID, 0) = ""
    SaveUser ID
    For t% = 0 To 50
        UserData(ID, t%) = ""
    Next t%
    PlayerLocHistory(ID) = String$(1500, 0)
    ConnInfo(ID, 1) = "0"
    Exit Sub
    Call WebWriteWhoInGame
    
End If

If Left$(Action$, 7) = "SAYFILE" Then
    Filename$ = " " + Trim$(Mid$(Action$, 8)) + " "
    i% = InStr(UCase$(Filename$), "$PARAM$")
    If i% <> 0 Then
        Filename$ = Left$(Filename$, i% - 1) + Param$ + Mid$(Filename$, i% + 7)
    End If
    Filename$ = Trim$(Filename$)
    SayFile ID, Filename$
    Exit Sub
End If

If Left$(Action$, 7) = "PORTS" Then
    WhoInGame ID, Param$
End If

If Action$ = "SYSPING" Then
    Log "SYSPING from " + Str$(ID) + " [" + UserData(ID, 0) + ": " + Param$ + "]"
    DesktopAlert "MUD32", "SYSPing from " + UserData(ID, 0) + ", '" + Param$ + "'"
    SendDDE "SYSPING from " + UserData(ID, 0) + ", '" + Param$ + "'"
    If Param$ = "" Then
        SayFile ID, "misc\sysping.txt"
    End If
    Beep
End If

If Left$(Action$, 4) = "LOOK" Then
    PlayerLoc = Val(UserData(ID, 10))
    LookAtLocation ID, PlayerLoc, 2
    WhosHere ID, PlayerLoc
End If

If Left$(Action$, 9) = "QUICKLOOK" Then
    PlayerLoc = Val(UserData(ID, 10))
    LookAtLocation ID, PlayerLoc, 1
    WhosHere ID, PlayerLoc
End If
If Left$(Action$, 4) = "TELL" Then
    SayToPerson ID, Param$
End If

If Left$(Action$, 3) = "ASK" Then
    AskPerson ID, Param$
End If

If Left$(Action$, 3) = "SAY" Then
    PlayerLoc = Val(UserData(ID, 10))
    SayToRoom ID, PlayerLoc, Param$
End If
    
If Left$(Action$, 3) = "ME" Then
    PlayerLoc = Val(UserData(ID, 10))
    EmoteToRoom ID, PlayerLoc, Param$
End If

If Left$(Action$, 6) = "SUFFIX" Then
     UserData(ID, 4) = Trim$(Left$(Param$, 15)) + " "
     If UserData(ID, 4) = " " Then UserData(ID, 4) = ""
     SendData ID, "New suffix '" + Trim$(UserData(ID, 4)) + "' set."
End If

If Left$(Action$, 7) = "PROFILE" Then
     UserData(ID, 6) = Trim$(Left$(Param$, 250))
     SendData ID, "New profile set. Type EXAM ME to see it"
End If

If Left$(Action$, 5) = "BRIEF" Then
    If UserData(ID, 5) = "" Then UserData(ID, 5) = "0"
    If UCase$(Trim$(Param$)) = "ON" Then UserData(ID, 5) = "1"
    If UCase$(Trim$(Param$)) = "OFF" Then UserData(ID, 5) = "0"
    
    If UserData(ID, 5) = "0" Then OutStr$ = "Brief Mode OFF. Locations will always be described in full."
    If UserData(ID, 5) = "1" Then OutStr$ = "Brief Mode ON. Locations already visited this session will not be described."
         
    SendData ID, OutStr$
End If
    
If Left$(Action$, 2) = "ID" Then
    SendData ID, "This is Location " + UserData(ID, 10)
    ExitList$ = ""
    For t% = 0 To 9
        If PlayerExit(ID, t%, 1) <> "" Then
            Z% = Len(ExitList$)
            ExitList$ = ExitList$ + ", " + PlayerExit(ID, t%, 1)
        End If
    Next t%
    If Z% <> 0 Then ExitList$ = Left$(ExitList$, Z%) + " and" + Mid$(ExitList$, Z% + 2)
    ExitList$ = Trim$(Mid$(ExitList$, 1))
    If ExitList$ <> "" Then ExitList$ = "Exits are" + Mid$(ExitList$, 2) + "." Else ExitList$ = "There are no exits."
    If Left$(UCase$(Param$), 1) = "E" Then SendData ID, ExitList$
    
End If

If Left$(Action$, 3) = "GET" Then
    'SendData ID, "GET Function"
    If Param$ = "" Then Param$ = "all"
    PickUpItem ID, Param$
End If

If Left$(Action$, 4) = "DROP" Then
    'SendData ID, "DROP Function"
    If Param$ = "" Then
        SendData ID, "You must specify what you want to drop."
    Else
        DropItem ID, Param$
    End If
End If

If Left$(Action$, 4) = "GIVE" Then
    'SendData ID, "GIVE Function"
    GiveItem ID, Param$
End If

If Left$(Action$, 5) = "STEAL" Then
    'SendData ID, "STEAL Function"
    StealItem ID, Param$
End If

If Left$(Action$, 9) = "INVENTORY" Then
    If Param$ = "" Then Param$ = UserData(ID, 0)
    PlayerInventory ID, Param$
End If

If Left$(Action$, 7) = "EXAMINE" Then
    If UCase$(Param$) = "ME" Then Param$ = UserData(ID, 0)
    Examine ID, Param$
End If

If Left$(Action$, 5) = "SCORE" Then
    ShowMyScore ID
End If

If Left$(Action$, 5) = "ADMIN" Then
    If Val(UserData(ID, 9)) > 2 Then
        ConsoleCommand ID, Param$
    Else
        SendData ID, "You do not have the required privileges to use this command."
    End If
End If


If Left$(Action$, 5) = "TOP" Then
    ShowTopTen ID
End If

If Left$(Action$, 6) = "LASTON" Then
    SayFile ID, "misc\laston.txt"
    SayFile ID, "misc\laston.dat"
End If

If Left$(Action$, 8) = "PASSWORD" Then
    i% = InStr(Param$, " ")
    If i% <> 0 Then
        OPass$ = Left$(Param$, i% - 1)
        NPass$ = Mid$(Param$, i% + 1)
        If InStr(NPass$, " ") <> 0 Then
            SendData ID, "Invalid characters in new password."
        Else
            Resp$ = ChangePlayerPass(UserData(ID, 0), OPass$, NPass$, 1)
            If Resp$ <> "" Then
                UserData(ID, 1) = Resp$
                Resp$ = "Your password has been changed."
            End If
            SendData ID, Resp$
        End If
    Else
        SendData ID, "Incorrect Parameters."
    End If
End If

If Left$(Action$, 3) = "CLR" Then
    ' fix me - build a set of VTTerm settings ?
    Text$ = Chr$(27) + "[;H" + Chr$(27) + "[2J"
    SendData ID, Text$
End If

If Left$(Action$, 8) = "TELEPORT" Then
    JoinPlayer ID, Param$
End If

If Left$(Action$, 4) = "DATE" Then
    Output$ = "It's " + Format$(Time, "Long time") + " " + ReadINI("MUD32.INI", "TimeZone") + " on " + Format$(Date, "Long date")
    SendData ID, Output$
End If

If Left$(Action$, 6) = "ATTACK" Then
    AttackPlayer ID, Param$
End If

If Left$(Action$, 9) = "RETALIATE" Then
    RetPlayer ID, Param$
End If

If Left$(Action$, 4) = "FLEE" Then
    Flee ID
End If
If Left$(Action$, 5) = "SHOUT" Then
    PlayerShout ID, Param$
End If
If Left$(Action$, 5) = "BUILD" Then
    Builder ID, Param$
End If

If Left$(Action$, 5) = "WRITE" Then
    WriteOnWall ID, Param$
End If

If Left$(Action$, 4) = "READ" Then
    ReadWall ID, Param$
End If

If Left$(Action$, 4) = "ECHO" Then
    If UCase$(Param$) = "OFF" Then
        UserData(ID, 7) = "1"
    End If
    If UCase$(Param$) = "ON" Then
        UserData(ID, 7) = "0"
    End If
    
    If Val(UserData(ID, 7)) = 1 Then tmp$ = "OFF" Else tmp$ = "ON"
    SendData ID, "ECHO is " + tmp$
    
End If

If Left$(Action$, 9) = "STATUSBAR" Then
    If UCase$(Param$) = "ON" Then
        UserData(ID, 8) = "1"
    End If
    If UCase$(Param$) = "OFF" Then
        UserData(ID, 8) = "0"
    End If
    
    If Val(UserData(ID, 8)) = 1 Then tmp$ = "ON" Else tmp$ = "OFF"
    SendData ID, "Status Bar is " + tmp$
    
End If

If Left$(Action$, 7) = "COLOUR" Then
    If UCase$(Param$) = "ON" Then
        UserData(ID, 21) = "1"
    End If
    If UCase$(Param$) = "OFF" Then
        UserData(ID, 21) = "0"
    End If
    
    If Val(UserData(ID, 21)) = 1 Then tmp$ = "ON" Else tmp$ = Chr$(27) + "[40;0;37mOFF"
    SendData ID, "Colour is " + tmp$

End If

If Left$(Action$, 6) = "UPTIME" Then
    SendData ID, ReadINI("mud32.ini", "WINTITLE") + " has been online for " + GetUptime
End If

If Left$(Action$, 6) = "WIDTH" Then

    If Param$ = "" Then
        If UserData(ID, 22) = "" Then UserData(ID, 22) = Str$(75)
        SendData ID, "Current terminal width is" & UserData(ID, 22) & "."
    Else
        uWidth = Val(Param$)
        If uWidth < 15 Or uWidth > 200 Then
            SendData ID, "Invalid terminal width value, please use a value between 15 and 200."
        Else
            UserData(ID, 22) = Str$(uWidth)
            SendData ID, "Terminal width is now" & UserData(ID, 22) & "."
        End If
    End If
    
End If

'******************************************************************************

End Sub

Sub LookAtLocation(ID, PlayerLoc, Flag)
LocFlags(ID) = ""
If UserData(ID, 5) <> "0" Then
    If GetBit(PlayerLoc, PlayerLocHistory(ID)) = True Then FlagHist = 1 Else FlagHist = 0
End If

'If GetBit(PlayerLoc, PlayerLocHistory(ID)) = True Then FlagHist = 1 Else FlagHist = 0

LocFlags(ID) = ""
LocFlags(ID) = GetLocFlags(PlayerLoc)
If InStr(LocFlags(ID), "+DARK") Then LocDark = 1
If InStr(UserData(ID, 17), "-DARK") Then LocDark = 0

Handle = FreeFile
Open App.Path + "\locations\" + Trim$(Str$(PlayerLoc)) + ".loc" For Input Access Read As Handle
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    LnType$ = Left$(Ln$, 1)
    Ln$ = Mid$(Ln$, 3)
    
    If LnType$ = "0" Then
        'Debug.Print VTBold(Val(UserData(ID, 8)), 1) + Ln$ + VTBold(Val(UserData(ID, 8)), 0)
        SendData ID, AnsiTable(Val(UserData(ID, 21)), 1) + Ln$ + AnsiTable(Val(UserData(ID, 21)), 0)
        UserData(ID, 20) = Ln$
    End If
    
    If LnType$ = "1" And LocDark <> 1 Then
        If FlagHist = 0 And Flag = 0 Then ShowLoc = 1
        If Flag = 2 Then ShowLoc = 1
        If ShowLoc = 1 Then
            If ParaFlag = 1 Then SendData ID, ""
            SendData ID, AnsiTable(Val(UserData(ID, 21)), 2) + Ln$ + AnsiTable(Val(UserData(ID, 21)), 0)
            ParaFlag = 1
            DontDoExits = 1
        End If
    End If
Loop

Close Handle

If LocDark = 1 Then
    SendData ID, AnsiTable(Val(UserData(ID, 21)), 2) + "It is too dark to see anything here." + AnsiTable(Val(UserData(ID, 21)), 0)
End If

'If UserData(ID, 5) = "2" And DontDoExits <> 1 Then

If ShowExits = 1 And LocDark <> 1 Then 'now controlled by a param in MUD32.ini
    ExitList$ = ""
    For t% = 0 To 9
        If PlayerExit(ID, t%, 1) <> "" Then
            Z% = Len(ExitList$)
            ExitList$ = ExitList$ + ", " + PlayerExit(ID, t%, 1)
        End If
    Next t%
    If Z% <> 0 Then ExitList$ = Left$(ExitList$, Z%) + " and" + Mid$(ExitList$, Z% + 2)
    ExitList$ = Trim$(Mid$(ExitList$, 1))
    If ExitList$ <> "" Then ExitList$ = "Exits are" + Mid$(ExitList$, 2) + "." Else ExitList$ = "There are no exits."
    SendData ID, AnsiTable(Val(UserData(ID, 21)), 1) + ExitList$ + AnsiTable(Val(UserData(ID, 21)), 0)
End If

ItemDesc$ = ""
For t% = 0 To 1000
    If ItemArray(t%, 0) = "1" Then
        If Val(ItemArray(t%, 1)) = PlayerLoc Then
            'LoadItemData ID, Trim$(Str$(T%)) + ".itx"
            ItemDesc$ = ReadINI("items\" + Trim$(Str$(t%)) + ".itx", "LocDesc")
            If LocDark <> 1 Or InStr(ItemArray(t%, 7), "-DARK") <> 0 Then
                SendData ID, AnsiTable(Val(UserData(ID, 21)), 8) + Trim$(ItemDesc$) + AnsiTable(Val(UserData(ID, 21)), 0)
            End If
        End If
    End If
Next t%
'If ItemDesc$ <> "" Then SendData ID, AnsiTable(Val(UserData(ID, 21)), 2) + Trim$(ItemDesc$) + AnsiTable(Val(UserData(ID, 21)), 0)

Call SetBit(PlayerLoc, PlayerLocHistory(ID), 1)

End Sub
Sub WhosHere(ID, PlayerLoc)
LocFlags(ID) = GetLocFlags(PlayerLoc)
If InStr(LocFlags(ID), "DARK") <> 0 Then Exit Sub

For t% = 0 To 64
    If Val(UserData(t%, 10)) = PlayerLoc And UserData(t%, 0) <> "" And t% <> ID Then
        SendData ID, UserData(t%, 0) + " the " + UserData(t%, 4) + PlayerLevel(t%, Val(UserData(t%, 3))) + " is here."
    End If
Next t%

NPC_LocLook ID, PlayerLoc

End Sub

Sub SayToRoom(ID, PlayerLoc, Mess$)
If Mess$ = "" Then
    SendData ID, "You clear your throat, as if to say something."
    Exit Sub
End If

SendData ID, "You say, " + Chr$(34) + Mess$ + Chr$(34)
If InStr(GetLocFlags(PlayerLoc), "DARK") = 0 Then
    Prefix$ = UserData(ID, 0)
Else
    Prefix$ = "Someone"
End If

For t% = 0 To 64
    If Val(UserData(t%, 10)) = PlayerLoc And t% <> ID Then
        SendData t%, Prefix$ + " says, " + Chr$(34) + Mess$ + Chr$(34)
        SayPrompt t%, ""

    End If
Next t%

End Sub
Sub PlayerShout(ID, Mess$)
If Mess$ = "" Then
    SendData ID, "You clear your throat, as if to say something."
    Exit Sub
End If


SendData ID, "You shout, " + Chr$(34) + Mess$ + Chr$(34)
For t% = 0 To 64
    If t% <> ID And ConnInfo(t%, 1) = "99" Then
        If Val(UserData(t%, 10)) = PlayerLoc Then
            SendData t%, UserData(ID, 0) + " Shouts, " + Chr$(34) + Mess$ + Chr$(34)
            SayPrompt t%, ""
        Else
            SendData t%, "A voice in the distance shouts, " + Chr$(34) + Mess$ + Chr$(34)
            SayPrompt t%, ""
        End If
    End If
Next t%

End Sub
Sub SayToPerson(ID, Mess$)
Mess$ = Trim$(Mess$)
i% = InStr(Mess$, " ")
If i% > 0 Then
    PlrName$ = Left$(Mess$, i% - 1)
    Mess$ = Mid$(Mess$, i% + 1)
Else
    SendData ID, "Tell WHO WHAT ?"
    Exit Sub
End If
If UCase$(UserData(ID, 0)) = UCase$(PlrName$) Then
    SendData ID, "Talking to yourself is a sign of madness, you know."
    Exit Sub
End If

For t% = 0 To 64
    If UserData(ID, 10) = UserData(t%, 10) Then
        If UCase$(UserData(t%, 0)) = UCase$(PlrName$) Then
            SendData ID, "You say to " + UserData(t%, 0) + ", " + Chr$(34) + Mess$ + Chr$(34)
            SendData t%, UserData(ID, 0) + " tells you, " + Chr$(34) + Mess$ + Chr$(34)
            SayPrompt t%, ""
            Who = t%
            t% = 100
        End If
    End If
Next t%


If t% < 100 Then
    'Maybe an NPC
    NPCName$ = NPC_IsHere(Val(UserData(ID, 10)), PlrName$)
    If NPCName$ <> "" Then
        NPC_TalkTo ID, NPCName$, Mess$
    Else
        SendData ID, "You don't see them here."
    End If
    Exit Sub
End If

For t% = 0 To 64
    If t% <> ID And t% <> Who Then
        If UserData(ID, 10) = UserData(t%, 10) Then
            SendData t%, UserData(ID, 0) + " talks to " + UserData(Who, 0) + "."
            SayPrompt t%, ""
        End If
    End If
    
Next t%

End Sub
Sub AskPerson(ID, Mess$)
Mess$ = Trim$(Mess$)
i% = InStr(Mess$, " ")
If i% > 0 Then
    PlrName$ = Left$(Mess$, i% - 1)
    Mess$ = Trim(Mid$(Mess$, i% + 1))
    If UCase$(Left$(Mess$, 5)) = "ABOUT" Then Mess$ = Trim(Mid$(Mess$, 6))
Else
    SendData ID, "Ask WHO about WHAT ?"
    Exit Sub
End If
If UCase$(UserData(ID, 0)) = UCase$(PlrName$) Then
    SendData ID, "Talking to yourself is a sign of madness, you know."
    Exit Sub
End If

For t% = 0 To 64
    If UserData(ID, 10) = UserData(t%, 10) Then
        If UCase$(UserData(t%, 0)) = UCase$(PlrName$) Then
            SendData ID, "You ask " + UserData(t%, 0) + ", " + Chr$(34) + Mess$ + Chr$(34)
            SendData t%, UserData(ID, 0) + " asks you, " + Chr$(34) + Mess$ + Chr$(34)
            SayPrompt t%, ""
            Who = t%
            t% = 100
        End If
    End If
Next t%


If t% < 100 Then
    'Maybe an NPC
    NPCName$ = NPC_IsHere(Val(UserData(ID, 10)), PlrName$)
    If NPCName$ <> "" Then
        NPC_TalkTo ID, NPCName$, Mess$
    Else
        SendData ID, "You don't see them here."
    End If
    Exit Sub
End If

For t% = 0 To 64
    If t% <> ID And t% <> Who Then
        If UserData(ID, 10) = UserData(t%, 10) Then
            SendData t%, UserData(ID, 0) + " asks " + UserData(Who, 0) + " a question."
            SayPrompt t%, ""
        End If
    End If
    
Next t%

End Sub
Sub EmoteToRoom(ID, PlayerLoc, Mess$)
SendData ID, UserData(ID, 0) + " " + Mess$
If InStr(GetLocFlags(PlayerLoc), "DARK") = 0 Then
    Prefix$ = UserData(ID, 0)
Else
    Prefix$ = "Someone"
End If
For t% = 0 To 64
    If Val(UserData(t%, 10)) = PlayerLoc And t% <> ID Then
        SendData t%, Prefix$ + " " + Mess$
        SayPrompt t%, ""

    End If
Next t%

End Sub
Sub SysEmoteToRoom(ID, PlayerLoc, Mess$)
'SendData ID, "You " + Mess$
If InStr(GetLocFlags(PlayerLoc), "DARK") = 0 Then
    Prefix$ = UserData(ID, 0) + " the " + UserData(ID, 4) + PlayerLevel(ID, UserData(ID, 3))
Else
    Prefix$ = "Someone"
End If

For t% = 0 To 64
    If Val(UserData(t%, 10)) = PlayerLoc And t% <> ID Then
        Text$ = "* " + Prefix$ + " " + Mess$
        SendData t%, Text$
        SayPrompt t%, ""

    End If
Next t%

End Sub

Sub Broadcast(Mess$)
For t% = 0 To 64
    If ConnInfo(t%, 1) = "99" Then
        SendData t%, Mess$
        SayPrompt t%, ""

    End If
Next t%


End Sub
Sub WhoInGame(ID, Param$)
    Header$ = Left$("Player Name" + Space$(50), 50) + "Status / Online Time"
    SendData ID, Header$
    Header$ = Left$("-----------" + Space$(50), 50) + "--------------------"
    SendData ID, Header$
    
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
            SendData ID, Txt$
        End If
    Next t%
    Footer$ = Trim$(Str$(ActivePorts)) + " Connected Players."
    SendData ID, String$(Len(Footer$), "-")
    SendData ID, Trim$(Str$(ActivePorts)) + " Connected Players."

End Sub

Sub PickUpItem(ID, Param$)
PickedUp = 0
PlayerLoc = Val(UserData(ID, 10))
For t% = 0 To 1000
    If ItemArray(t%, 0) = "1" Then
        If Val(ItemArray(t%, 1)) = PlayerLoc Then
            If InStr(UCase(ItemArray(t%, 2)), UCase$(Param$)) <> 0 Or UCase$(Param$) = "ALL" Then
                CurWeight = Val(UserData(ID, 16))
                If CurWeight + Val(ItemArray(t%, 4)) > Val(UserData(ID, 15)) Then
                    ItemFile$ = "items\" + Trim$(Str$(t%)) + ".itx"
                    PickUpFail$ = ReadINI(ItemFile$, "tooheavy")
                    If PickUpFail$ = "" Then PickUpFail$ = "You cannot pick up " + ItemArray(t%, 3) + ", You are carrying too much."
                    SendData ID, PickUpFail$
                    PickedUp = 1
                Else
                    SendData ID, "You pickup " + ItemArray(t%, 3) + "."
                    SysEmoteToRoom ID, PlayerLoc, "picks up " + ItemArray(t%, 3) + "."
                    ItemArray(t%, 0) = Trim$(Str$(2))
                    ItemArray(t%, 1) = Trim$(Str$(ID))
                    CurWeight = CurWeight + Val(ItemArray(t%, 4))
                    UserData(ID, 16) = Trim$(Str$(CurWeight))
                    ItemFlag_Add ID, t%
                    PickedUp = 1
                    If UCase$(Param$) <> "ALL" Then t% = 1001
                End If
            End If
        End If
    End If
Next t%
If PickedUp = 0 Then SendData ID, "You don't see that here."

End Sub
Sub DropItem(ID, Param$)
Dropped = 0
PlayerLoc = Val(UserData(ID, 10))
For t% = 0 To 1000
    If ItemArray(t%, 0) = "2" Then
        If Val(ItemArray(t%, 1)) = ID Then
            If InStr(UCase(ItemArray(t%, 2)), UCase$(Param$)) <> 0 Or UCase$(Param$) = "ALL" Then
                SendData ID, "You drop " + ItemArray(t%, 3) + "."
                DropVal = DropVal + Val(ItemArray(t%, 5))
                SysEmoteToRoom ID, PlayerLoc, "drops " + ItemArray(t%, 3) + "."
                If PlayerLoc = ScoreLoc And Val(ItemArray(t%, 5)) <> 0 Then
                    ItemArray(t%, 0) = Trim$(Str$(0))
                    AddToRList t%
                Else
                    ItemArray(t%, 0) = Trim$(Str$(1))
                End If
                ItemArray(t%, 1) = Trim$(Str$(PlayerLoc))
                CurWeight = Val(UserData(ID, 16))
                CurWeight = CurWeight - Val(ItemArray(t%, 4))
                UserData(ID, 16) = Trim$(Str$(CurWeight))
                ItemFlag_Remove ID, t%
                Dropped = 1
                If UCase$(Param$) <> "ALL" Then t% = 1001
            End If
        End If
    End If
Next t%
If Dropped = 0 And UCase$(Param$) <> "ALL" Then SendData ID, "You are not carrying that."
If PlayerLoc = ScoreLoc And DropVal <> 0 Then
    SayFile ID, "misc\score.txt"
    OldLevel = PlayerLevel(ID, UserData(ID, 3))
    MaxHealth = Levels(PlayerLevelNo(UserData(ID, 3)) - 1, 3)
    Health = Val(UserData(ID, 11))
    If Health < MaxHealth Then
        Health = Health + DropVal
        If Health > MaxHealth Then
            DropVal = Health - MaxHealth
            Health = MaxHealth
        Else
            DropVal = 0
        End If
        UserData(ID, 11) = Trim$(Str$(Health))
        SendData ID, "Your health has increased and you feel refreshed."
    End If
        
    If DropVal > 0 Then
        Score = Val(UserData(ID, 3))
        Score = Score + DropVal
        UserData(ID, 3) = Trim$(Str$(Score))
        SendData ID, "Your score is now " + Format$(Val(UserData(ID, 3)), "###,###,##0") + "."
    End If
    
    UserData(ID, 15) = Levels(PlayerLevelNo(UserData(ID, 3)) - 1, 3)
    If UserData(ID, 11) = MaxHealth Then UserData(ID, 11) = Levels(PlayerLevelNo(UserData(ID, 3)) - 1, 3)
    
    NewLevel = PlayerLevel(ID, UserData(ID, 3))
    Call UpdateTopTen(UserData(ID, 0), UserData(ID, 3), UserData(ID, 2))
    If OldLevel <> NewLevel Then
        Text$ = "Congratulations! You are now " + UserData(ID, 0) + " the " + UserData(ID, 4) + PlayerLevel(ID, Val(UserData(ID, 3))) + "!"
        SendData ID, Text$
        SendDDE UserData(ID, 0) + " has attained the level of " + PlayerLevel(ID, Val(UserData(ID, 3))) + "!"
    End If
    
    'Are we immortal yet ?
    
    If PlayerLevelNo(UserData(ID, 3)) >= GodLevel Then
        If Val(UserData(ID, 9)) <> 3 Then SayFile ID, "misc\immortal.txt"
        UserData(ID, 9) = 3
    End If
    
End If
End Sub


Sub PlayerInventory(ID, Param$)
For t% = 0 To 64
    If UCase$(UserData(t%, 0)) = UCase$(Param$) Then
        Who = t%: Found = 1
        t% = 65
    End If
Next t%
If UserData(Who, 10) <> UserData(ID, 10) Or Found = 0 Then
    SendData ID, "You don't see them here."
    Exit Sub
End If
If Who = ID Then Prefix$ = "You are " Else Prefix$ = UserData(Who, 0) + " is "
items = 0
SendData ID, Prefix$ + "carrying:"
For t% = 0 To 1000
    If ItemArray(t%, 0) = "2" Then 'item is 'in hand'
        If Val(ItemArray(t%, 1)) = Who Then 'specified player is carrying it
            SendData ID, "  " + ItemArray(t%, 3) '+ " (" + ItemArray(T%, 4) + ")"
            items = 1
        End If
    End If
Next t%
If items = 0 Then SendData ID, "  nothing"

End Sub
Sub GiveItem(ID, Param$)
If Param$ = "" Then
    SendData ID, "Give what to who ?"
    Exit Sub
End If
i% = InStr(UCase$(Param$), " TO ")
If i% <> 0 Then Param$ = Left$(Param$, i% - 1) + " " + Mid$(Param$, i% + 4)
i% = InStr(UCase$(Param$), " ")
If i% <> 0 Then
    ItemName$ = Trim$(Left$(Param$, i% - 1))
    Player$ = Trim$(Mid$(Param$, i% + 1))
End If
Param$ = ItemName$

For t% = 0 To 64
    If UCase$(UserData(t%, 0)) = UCase$(Player$) Then
        Who = t%: Found = 1
        t% = 65
    End If
Next t%
If Who = ID And Found = 1 Then
    SendData ID, "Give to yourself? Don't be silly."
    Exit Sub
End If

If Found = 0 Then
    SendData ID, "You don't see them here."
    Exit Sub
End If
If Found = 1 And UserData(Who, 10) <> UserData(ID, 10) Then
    SendData ID, "You don't see them nearby."
    Exit Sub
End If

Given = 0
PlayerLoc = Val(UserData(ID, 10))
For t% = 0 To 1000
    If ItemArray(t%, 0) = "2" Then
        If Val(ItemArray(t%, 1)) = ID Then
            If InStr(UCase(ItemArray(t%, 2)), UCase$(Param$)) <> 0 Or UCase$(Param$) = "ALL" Then
                CurWeight = Val(UserData(Who, 16))
                If CurWeight + Val(ItemArray(t%, 4)) > Val(UserData(Who, 15)) Then
                    SendData ID, "You try to give " + ItemArray(t%, 3) + " to " + UserData(Who, 0) + " but fail."
                    Given = 1
                Else
                    SendData ID, "You give " + ItemArray(t%, 3) + " to " + UserData(Who, 0) + "."
                    SendData Who, "* " + UserData(ID, 0) + " gives you " + ItemArray(t%, 3) + "."
                    SayPrompt Who, ""
                    ItemArray(t%, 1) = Trim$(Str$(Who))
                    myWeight = Val(UserData(ID, 16))
                    myWeight = myWeight - Val(ItemArray(t%, 4))
                    UserData(ID, 16) = Trim$(Str$(myWeight))
                    ItemFlag_Remove ID, t%
                    Given = 1
                    CurWeight = CurWeight + Val(ItemArray(t%, 4))
                    UserData(Who, 16) = Trim$(Str$(CurWeight))
                    ItemFlag_Add Who, t%
                    If UCase$(Param$) <> "ALL" Then t% = 1001
                End If
            End If
        End If
    End If
Next t%
If Given = 0 And UCase$(Param$) <> "ALL" Then SendData ID, "You are not carrying that."
If Given = 0 And UCase$(Param$) = "ALL" Then SendData ID, "You have nothing to give."

End Sub
Sub StealItem(ID, Param$)
If InStr(LocFlags(ID), "SAFE") <> 0 Then
    SendData ID, "You cannot steal from someone here. This is a safe zone."
    Exit Sub
End If
If Param$ = "" Then
    SendData ID, "Steal what from who ?"
    Exit Sub
End If
i% = InStr(UCase$(Param$), " FROM ")
If i% <> 0 Then Param$ = Left$(Param$, i% - 1) + " " + Mid$(Param$, i% + 6)
i% = InStr(UCase$(Param$), " ")
If i% <> 0 Then
    ItemName$ = Trim$(Left$(Param$, i% - 1))
    Player$ = Trim$(Mid$(Param$, i% + 1))
End If
Param$ = ItemName$

For t% = 0 To 64
    If UCase$(UserData(t%, 0)) = UCase$(Player$) Then
        Who = t%: Found = 1
        t% = 65
    End If
Next t%
If Who = ID And Found = 1 Then
    SendData ID, "Steal from yourself? Don't be silly."
    Exit Sub
End If

If Found = 0 Then
    SendData ID, "You don't see them here."
    Exit Sub
End If
If Found = 1 And UserData(Who, 10) <> UserData(ID, 10) Then
    SendData ID, "You don't see them nearby."
    Exit Sub
End If

Taken = 0
PlayerLoc = Val(UserData(ID, 10))
For t% = 0 To 1000
    If ItemArray(t%, 0) = "2" Then
        If Val(ItemArray(t%, 1)) = Who Then
            If InStr(UCase(ItemArray(t%, 2)), UCase$(Param$)) <> 0 Then
                CurWeight = Val(UserData(ID, 16))
                If CurWeight + Val(ItemArray(t%, 4)) > Val(UserData(ID, 15)) Then
                    SendData ID, "You cannot steal that, as you are carrying to much already"
                    'SendData ID, "You try to steal " + ItemArray(T%, 3) + " from " + UserData(Who, 0) + " but fail."
                    'SendData Who, "* " + UserData(ID, 0) + " tried to steal " + ItemArray(T%, 3) + " from you, but failed."
                    'SayPrompt Who, ""
                    Taken = 1
                ElseIf Success(ID, 40, Val(UserData(ID, 3))) = False Then
                    SendData ID, "You try to steal " + ItemArray(t%, 3) + " from " + UserData(Who, 0) + " but fumble it badly!"
                    SendData Who, "* " + UserData(ID, 0) + " tried to steal " + ItemArray(t%, 3) + " from you, but failed."
                    SayPrompt Who, ""
                    Taken = 1
                Else
                    SendData ID, "You steal " + ItemArray(t%, 3) + " from " + UserData(Who, 0) + "."
                    SendData Who, "* " + UserData(ID, 0) + " has stolen " + ItemArray(t%, 3) + " from you!"
                    SayPrompt Who, ""
                    ItemArray(t%, 1) = Trim$(Str$(ID))
                    oldWeight = Val(UserData(Who, 16))
                    oldWeight = oldWeight - Val(ItemArray(t%, 4))
                    UserData(Who, 16) = Trim$(Str$(oldWeight))
                    Taken = 1
                    CurWeight = CurWeight + Val(ItemArray(t%, 4))
                    UserData(ID, 16) = Trim$(Str$(CurWeight))
                    ItemFlag_Add ID, t%
                    ItemFlag_Remove Who, t%
                    If UCase$(Param$) <> "ALL" Then t% = 1001
                End If
            End If
        End If
    End If
Next t%
If Taken = 0 Then SendData ID, UserData(Who, 0) + " is not carrying that."

End Sub

Sub Examine(ID, Param$)

'are we looking for a player?

For t% = 0 To 64
    If UCase$(UserData(t%, 0)) = UCase$(Param$) Then
        Who = t%: Found = 1
        t% = 65
    End If
Next t%
If Found = 1 Then 'Param is a Player
    If UserData(Who, 10) <> UserData(ID, 10) Then
        SendData ID, "You don't see them nearby."
        Exit Sub
    End If
    DescribePlayer ID, Who
    Exit Sub
End If

'are we looking for an NPC?
Resp$ = NPC_Examine(UserData(ID, 10), Param$)
If Resp$ <> "" Then
    SendData ID, Resp$
    Exit Sub
End If

'are we looking for an item?

For t% = 0 To 1000
    If ItemArray(t%, 0) = "2" Then
        If Val(ItemArray(t%, 1)) = ID Then
            If InStr(UCase(ItemArray(t%, 2)), UCase$(Param$)) <> 0 Then
                DescribeObject ID, t%
                Exit Sub
                t% = 1001
            End If
        End If
    End If
Next t%
SendData ID, "You are not carrying that."

End Sub

Sub DescribePlayer(ID, Who)
SendData ID, UserData(Who, 0) + " the " + UserData(Who, 4) + PlayerLevel(Who, UserData(Who, 3))
If UserData(Who, 6) <> "" Then SendData ID, UserData(Who, 6)
If UserData(Who, 2) = "M" Then
    Gender$ = "He has"
Else
    Gender$ = "She has"
End If
If ID = Who Then Gender$ = "You have"
SendData ID, Gender$ + " " + Trim(UserData(Who, 11)) + " health and " + UserData(Who, 15) + " strength. "
If Val(UserData(Who, 16)) > 0 Then SendData ID, "Items carried have " + UserData(Who, 16) + " total weight."
If Val(UserData(Who, 9)) <> 0 Then SendData ID, UserData(Who, 0) + " currently holds '" + SetUserLevel(Val(UserData(Who, 9))) + "' status."
End Sub

Sub DescribeObject(ID, Who)
ItemFile$ = "items\" + Trim$(Str$(Who)) + ".itx"
FullDesc$ = ReadINI(ItemFile$, "desc")
'SendData ID, ItemArray(Who, 3)
SendData ID, FullDesc$
tmp$ = "It has a value of " + ItemArray(Who, 5)
If Val(ItemArray(Who, 6)) <> 0 Then
    tmp$ = tmp$ + ", a weight of " + ItemArray(Who, 4) + " and can inflict " + ItemArray(Who, 6) + " damage."
Else
    tmp$ = tmp$ + " and a weight of " + ItemArray(Who, 4) + "."
End If
SendData ID, tmp$
End Sub

Sub ShowMyScore(ID)
SendData ID, "---------------------------" & UCase$(UserData(ID, 0)) & "--------------------------------"
SendData ID, "Race:  "
SendData ID, "Class: "
SendData ID, "Level: " & PlayerLevel(ID, UserData(ID, 3))
SendData ID, "Hp:    " & UserData(ID, 11)
SendData ID, "Exp:   " & Format$(Val(UserData(ID, 3)), "###,###,##0")
SendData ID, "Str:   " & UserData(ID, 15)
SendData ID, "Int:   "
SendData ID, "Dex    "
SendData ID, "Con    "
SendData ID, "Wis    "
SendData ID, "Char   "
SendData ID, "-------------------------------------------------------------------------------------------"
End Sub

Sub SysMsgToRoom(PlayerLoc, Mess$)
For t% = 0 To 64
    If Val(UserData(t%, 10)) = PlayerLoc Then
        Text$ = "* " + Mess$
        SendData t%, Text$
        SayPrompt t%, ""
    End If
Next t%

End Sub
Sub BugPrint(Mess$)
Handle = FreeFile
Open App.Path + "\mud32.dbg" For Output Access Write As Handle
Print #Handle, Date$ + " " + Time$
Print #Handle, Mess$
Close #Handle

End Sub
Sub ShowTopTen(ID)

SendData ID, "Pos  Name                    Score  Level             Last Played"
SendData ID, "-----------------------------------------------------------------"
Handle = FreeFile
Open App.Path + "\misc\top10.lst" For Input Access Read As #Handle
For t% = 1 To 10
    Input #Handle, Name$, Score$, Gender$, Played$
    If Gender$ = "M" Then Offset = -1 Else Offset = -2
    Level$ = PlayerLevel(Offset, Val(Score$))
    Ln$ = Right$("  " + Trim$(Str$(t%)), 2) + "   " + Left$(Name$ + Space$(20), 17) + Right$(Space$(12) + Format$(Val(Score$), "###,###,##0"), 12) + "  " + Left$(Level$ + Space$(20), 18) + Played$
    If Name$ <> "" Then SendData ID, Ln$
Next t%
Close #Handle
SendData ID, "-----------------------------------------------------------------"
End Sub
Sub TelePortPlayer(ID, Dest)
    SysEmoteToRoom ID, Val(UserData(ID, 10)), "teleports out."

    UserData(ID, 10) = Dest
    
    PlayerLoc = Val(UserData(ID, 10))
    
    GetExits ID, PlayerLoc
    SysEmoteToRoom ID, PlayerLoc, "teleports in."
                  
    LookAtLocation ID, PlayerLoc, 0
                
    WhosHere ID, PlayerLoc
End Sub

Sub JoinPlayer(ID, PlayerName$)
If PlayerName$ = "" Then
    SendData ID, "You must specify who you wish to join."
    Exit Sub
End If


SayFile ID, "misc\spell.txt"
If Success(ID, 0, Val(ItemData(ID, 3))) = True Then
    For t% = 0 To 64
        If UCase$(UserData(t%, 0)) = UCase$(PlayerName$) Then
            DestLoc = Val(UserData(t%, 10))
            t% = 65
        End If
    Next t%
        If DestLoc = 0 Then
        SendData ID, "That person is not in the world."
    Else
        TelePortPlayer ID, DestLoc
    End If
Else
    SayFile ID, "misc\spellfail.txt"
End If

End Sub

Sub AttackPlayer(ID, Param$)
If InStr(LocFlags(ID), "SAFE") <> 0 Then
    SendData ID, "You cannot attack someone here. This is a safe zone."
    Exit Sub
End If

If Val(UserData(ID, 9)) = 3 Then
    SendData ID, "Immortals cannot attack other players."
    Exit Sub
End If

If InStr(UCase$(Param$), "WITH") = 0 Then
    SendData ID, "Attack WHO with WHAT ?"
    Exit Sub
End If
'Parse Line, get Player$ and ItemName$
i% = InStr(UCase$(Param$), " WITH ")
If i% <> 0 Then Param$ = Left$(Param$, i% - 1) + " " + Mid$(Param$, i% + 6)
i% = InStr(UCase$(Param$), " ")
If i% <> 0 Then
    Player$ = Trim$(Left$(Param$, i% - 1))
    ItemName$ = Trim$(Mid$(Param$, i% + 1))
End If
Param$ = ItemName$

'Look for victim - player
For t% = 0 To 64
    If UCase$(UserData(t%, 0)) = UCase$(Player$) Then
        Who = t%: Found = 1: NPCAttack = 0
        t% = 65
    End If
Next t%
'look for victim - NPC
Resp$ = NPC_IsHere(UserData(ID, 10), Player$)
If Resp$ <> "" Then
    'we are attacking an NPC and it is here.
    NPCAttack = 1: NPCID = NPC_ID(Resp$)

End If


'If it's yourself be sarcastic ;-)
If Who = ID And Found = 1 Then
    SendData ID, "Attack yourself? Don't be silly."
    Exit Sub
End If

'If player not even in game, be suble with response (Astute players may suss this)
If Found = 0 And NPCAttack = 0 Then
    SendData ID, "You don't see them here."
    Exit Sub
End If

'If player found, but not in location inform accordingly
If Found = 1 And UserData(Who, 10) <> UserData(ID, 10) Then
    SendData ID, "You don't see them nearby."
    Exit Sub
End If

'check to see if victim is up for a squirmish
If Found = 1 And UserTemp(Who) <> "" Then
    If Right$(UserTemp(Who), 1) <> Chr$(ID) Then
        SendData ID, "They are already in a fight with someone else."
        Exit Sub
    End If
End If

'Check if victim is immortal
If Found = 1 And Val(UserData(Who, 9)) = 3 Then
    SendData ID, "Attack an immortal! Are you mad ?"
    Exit Sub
End If

'if the server has idle protect on, then prevent attacks on players who are AFK.
If Found = 1 And IdleProtect = 1 Then
    IdleTime = (Timer - Val(UserData(Who, 12))) / 60
    If IdleTime < 0 Then IdleTime = 0
    If Int(IdleTime) <> 0 Then
        SendData ID, "You try to attack them, but they don't seem to be paying attention."
        Exit Sub
    End If
End If

'If balancefights is set, then check level diff and prevent big players beating up little players.
IDLevel = PlayerLevelNo(Val(UserData(ID, 3)))
WhoLevel = PlayerLevelNo(Val(UserData(Who, 3)))

If Found = 1 Then

    If Abs(IDLevel - WhoLevel) > BalanceFights Then
        SendData ID, "You can't attack them. Pick on someone your own size."
        Exit Sub
    End If
    
End If

PlayerLoc = Val(UserData(ID, 10))

'Check weapon
For t% = 0 To 1000
    If ItemArray(t%, 0) = "2" Then
        If Val(ItemArray(t%, 1)) = ID Then
            If InStr(UCase(ItemArray(t%, 2)), UCase$(Param$)) <> 0 Then
                FoundWeap = 1
                ItemNo = t%
                t% = 1001
            End If
        End If
    End If
Next t%
If FoundWeap = 0 Then
    SendData ID, "You are not carrying that."
    Exit Sub
End If

Damage = Val(ItemArray(ItemNo, 6))
WeapName = ItemArray(ItemNo, 2)

If Damage < 1 Then
    SendData ID, "You can't attack someone with " + ItemArray(ItemNo, 3) + "!"
    Exit Sub
End If

'if we're not fighting another player,  let the NPC fight routine handle it.

If NPCAttack = 1 Then
    If NPCInfo(NPCID, 21) = "1" Then
        SendData ID, "You consider attacking " + NPCInfo(NPCID, 9) + ", but then think better of it."
        Exit Sub
    End If
    If NPCInfo(NPCID, 13) = "3" Then
        SendData ID, "You can't attack something that's already dead!"
    ElseIf NPCInfo(NPCID, 13) <> "0" And NPCID <> Val(UserData(ID, 14)) - 1 Then
        SendData ID, NPCInfo(NPCID, 9) & " is already in fight with someone else, wait your turn!"
    Else
        NPC_Attack ID, ItemArray(ItemNo, 3), Damage, NPCID
    End If
    Exit Sub
End If


If UserTemp(ID) = "" Then
    UserTemp(ID) = "ATTACK" + Chr$(Who)
    UserTemp(Who) = "VICTIM" + Chr$(ID)
End If

If Success(ID, 50, Val(UserData(ID, 3))) = True Then
    DrewBlood = 1
    SendData ID, "You attack " + UserData(Who, 0) + " with " + ItemArray(ItemNo, 3) + " and do" + Str$(Damage) + " damage!"
    SendData Who, "* " + UserData(ID, 0) + " sucessfully attacks you with " + ItemArray(ItemNo, 3) + "!"
    FightMsgToRoom PlayerLoc, UserData(ID, 0) + " attacks " + UserData(Who, 0) + " with " + ItemArray(ItemNo, 3) + "!"
    Health = Val(UserData(Who, 11))
    Health = Health - Damage: If Health < 0 Then Health = 0
    UserData(Who, 11) = Trim$(Str$(Health))
    SendData Who, "* Your Health is now " + UserData(Who, 11)
Else
    DrewBlood = 0
    SendData ID, "You attack " + UserData(Who, 0) + " with " + ItemArray(ItemNo, 3) + ", but miss!"
    SendData Who, "* " + UserData(ID, 0) + " tries to attack you with " + ItemArray(ItemNo, 3) + ", but misses!"
    FightMsgToRoom PlayerLoc, UserData(ID, 0) + " attacks " + UserData(Who, 0) + " with " + ItemArray(ItemNo, 3) + ", but misses!"
End If

If Health = 0 And DrewBlood = 1 Then
    SayFile Who, "misc\death.txt"
    SendData ID, "* " + UserData(Who, 0) + " has DIED. You Won!"
    
    For t% = 0 To 64
    If UserData(t%, 0) <> "" And t% <> ID And t% <> Who Then
        SendData t%, "* " + UserData(Who, 0) + " the " + UserData(Who, 4) + PlayerLevel(Who, UserData(Who, 3)) + " has been killed by " + UserData(ID, 0) + " the " + UserData(ID, 4) + PlayerLevel(ID, UserData(ID, 3)) + "!"
        SayPrompt t%, ""
    End If
    Next t%
    
    SendDDE UserData(Who, 0) + " has been killed by " + UserData(ID, 0) + "!"
    
    
    ' Sort out dead persons score.
       
    OldLevel = PlayerLevel(Who, UserData(Who, 3))
    Score = Val(UserData(Who, 3))
    VictorsPoints = CLng(Score / 2)
    If Left$(UserTemp(Who), 6) = "VICTIM" Then
        Score = CLng(Score / 2)
    Else
        Score = 0
    End If
    If Score < 0 Then Score = 0
    UserData(Who, 3) = Trim$(Str$(Score))
    SendData Who, "Your score is now " + Format$(Val(UserData(Who, 3)), "###,###,##0") + "."
    NewLevel = PlayerLevel(Who, UserData(Who, 3))
    Call UpdateTopTen(UserData(Who, 0), UserData(Who, 3), UserData(Who, 2))
    If OldLevel <> NewLevel Then
        Text$ = "You are now " + UserData(Who, 0) + " the " + UserData(Who, 4) + PlayerLevel(Who, Val(UserData(Who, 3))) + "."
        SendData Who, Text$
    End If
    SendData Who, ""
    
    ' Sort out Alive Persons score
    
    OldLevel = PlayerLevel(ID, UserData(ID, 3))
    Score = Val(UserData(ID, 3))
    Score = Score + VictorsPoints
    UserData(ID, 3) = Trim$(Str$(Score))
    SendData ID, "Your score is now " + Format$(Val(UserData(ID, 3)), "###,###,##0") + "."
    NewLevel = PlayerLevel(ID, UserData(ID, 3))
    Call UpdateTopTen(UserData(ID, 0), UserData(ID, 3), UserData(ID, 2))
    
    
    If OldLevel <> NewLevel Then
        Text$ = "Congratulations, you are now " + UserData(ID, 0) + " the " + UserData(ID, 4) + PlayerLevel(ID, Val(UserData(ID, 3))) + "!"
        SendData ID, Text$
        
        UserData(ID, 15) = Levels(PlayerLevelNo(UserData(ID, 3)) - 1, 3)
        If UserData(ID, 11) = MaxHealth Then UserData(ID, 11) = Levels(PlayerLevelNo(UserData(ID, 3)) - 1, 3)

    End If
    
    ReincarnatePlayer Who
    
    UserTemp(ID) = ""
    UserTemp(Who) = ""
 
    
    
End If

SayPrompt Who, ""

End Sub

Sub ReincarnatePlayer(ID)
    
    'drop all
    PlayerLoc = Val(UserData(ID, 10))
    For t% = 0 To 1000
    If ItemArray(t%, 0) = "2" Then
        If Val(ItemArray(t%, 1)) = ID Then
                ItemArray(t%, 0) = Trim$(Str$(1))
                ItemArray(t%, 1) = Trim$(Str$(PlayerLoc))
                CurWeight = Val(UserData(ID, 16))
                CurWeight = CurWeight - Val(ItemArray(t%, 4))
                UserData(ID, 16) = Trim$(Str$(CurWeight))
        End If
    End If
    Next t%
    
       
    Call ResetPlayerVars(ID, 1)
    
    PlayerLoc = Val(UserData(ID, 10))
    
    GetExits ID, PlayerLoc
    For t% = 0 To 64
    If Val(UserData(t%, 10)) = PlayerLoc And t% <> ID Then
        Text$ = "* " + "There is a blinding flash of light and " + UserData(ID, 0) + " the " + UserData(ID, 4) + PlayerLevel(ID, Val(UserData(ID, 3))) + " is reincarnated before you."
        SendData t%, Text$
        SayPrompt t%, ""
    End If
    Next t%
    
    SayFile ID, "misc\reincarnate.txt"
              
    LookAtLocation ID, PlayerLoc, 0
                
    WhosHere ID, PlayerLoc
End Sub

Sub Flee(ID)
If Left$(UserTemp(ID), 6) = "VICTIM" Or Left$(UserTemp(ID), 6) = "ATTACK" Then OkFlee = 1

If OkFlee = 0 Then
    SendData ID, "You can only flee during a fight."
    Exit Sub
End If

PlayerLoc = Val(UserData(ID, 10))
SysEmoteToRoom ID, PlayerLoc, "flees from the fight!"

UserTemp(ID) = ""
IsNPC = UserData(ID, 14) - 1
If IsNPC < 0 Then
    Who = Asc(Right$(UserTemp(ID), 1))
    UserTemp(Who) = ""
Else
    NPCInfo(IsNPC, 13) = "0"
    NPCInfo(IsNPC, 15) = "0"
    UserData(ID, 14) = 0
End If

OldLevel = PlayerLevel(ID, UserData(ID, 3))
Score = Val(UserData(ID, 3))
Score = Int(Score * 0.9)
If Score < 0 Then Score = 0
UserData(ID, 3) = Trim$(Str$(Score))
SendData ID, "Your score is now " + Format$(Val(UserData(ID, 3)), "###,###,##0") + "."
NewLevel = PlayerLevel(ID, UserData(ID, 3))
Call UpdateTopTen(UserData(ID, 0), UserData(ID, 3), UserData(ID, 2))
If OldLevel <> NewLevel Then
    Text$ = "You are now " + UserData(ID, 0) + " the " + UserData(ID, 4) + PlayerLevel(ID, Val(UserData(ID, 3))) + "."
    SendData ID, Text$
End If

For t% = 0 To 9
    If PlayerExit(ID, t%, 0) = "" Then NoMore = t%: t% = 10
Next t%

i% = Int(NoMore * Rnd)

UserData(ID, 10) = PlayerExit(ID, i%, 2)
PlayerLoc = Val(UserData(ID, 10))

GetExits ID, PlayerLoc
SysEmoteToRoom ID, PlayerLoc, "arrives."
                  
LookAtLocation ID, PlayerLoc, 0
                
WhosHere ID, PlayerLoc

'GetAction ID, Trim$(UCase$(PlayerExit(ID, I%, 0))), ""
Debug.Print PlayerExit(ID, i%, 0)

End Sub


Sub FightMsgToRoom(PlayerLoc, Mess$)
For t% = 0 To 64
    If Val(UserData(t%, 10)) = PlayerLoc And UserTemp(t%) = "" Then
        Text$ = "* " + Mess$
        SendData t%, Text$
        SayPrompt t%, ""
    End If
Next t%

End Sub

Sub RetPlayer(ID, Param$)
If UserTemp(ID) = "" Then
    SendData ID, "You can only retaliate during a fight, to start a fight, type ATTACK <player> WITH <weapon>"
    Exit Sub
End If

Who = Asc(Right$(UserTemp(ID), 1))
If Who > 0 Then
    WhoName = UserData(Who, 0)
Else
    'Grab NPC Name
    WhoName = Left$(UserTemp(ID), Len(UserTemp(ID)) - 1)
    WhoName = Mid$(WhoName, 7)
End If

Param$ = WhoName + " WITH " + Param$

Call AttackPlayer(ID, Param$)

End Sub

Sub WriteOnWall(ID, Param$)

If NoticeBoardLoc <> Val(UserData(ID, 10)) Then
    SendData ID, "Graffiti Artists are not tolerated here, Sorry."
    Exit Sub
End If
If Param$ = "" Then
    SendData ID, "Your hand hovers, but you are unsure what to write..."
    Exit Sub
End If

Dim WallBuffer(10) As String

SendData ID, "You write your message on the notice board."

Handle = FreeFile
Counter = 0
If IfExist(App.Path + "\misc\board.dat") Then
    Open App.Path + "\misc\board.dat" For Input Access Read As #Handle
    Do Until EOF(Handle)
        Line Input #Handle, WallBuffer(Counter)
        Counter = Counter + 1
    Loop
    Close Handle
End If

Handle = FreeFile
Open App.Path + "\misc\board.dat" For Output Access Write As #Handle
LowCounter = Counter - 9: If LowCounter < 0 Then LowCounter = 0
For t% = LowCounter To Counter - 1
    If WallBuffer(t%) <> "" Then Print #Handle, WallBuffer(t%)
Next t%
Print #Handle, "'" + Param$ + "' - " + UserData(ID, 0)
Close Handle

End Sub
Sub ReadWall(ID, Param$)
If NoticeBoardLoc <> Val(UserData(ID, 10)) Then
    SendData ID, "You cannot see the notice board from here!"
    Exit Sub
End If
SayFile ID, "misc\board.txt"
Handle = FreeFile
If IfExist(App.Path + "\misc\board.dat") Then
    Open App.Path + "\misc\board.dat" For Input Access Read As #Handle
    Do Until EOF(Handle)
        Line Input #Handle, Ln$
        SendData ID, Ln$
        Counter = Counter + 1
    Loop
    Close Handle
End If

End Sub

Sub BellTolls()
CMin$ = Mid$(Time$, 4, 2)
If CMin$ = "00" Then
    CHour% = Val(Left$(Time$, 2))
    CHour% = CHour% Mod 12
    If CHour% = 0 Then CHour% = 12
    If CHour% = 1 Then Amount$ = "once"
    If CHour% = 2 Then Amount$ = "twice"
    If CHour% = 3 Then Amount$ = "thrice"
    If CHour% > 3 Then Amount$ = Trim$(Str$(CHour%)) + " times"
    Broadcast "A bell tolls " + Amount$ + " faintly in the distance."
End If

End Sub

Sub ItemFlag_Add(ID, ItemNo)
Dim ArrayTmp() As String
ArrayTmp() = Split(ItemArray(ItemNo, 7), " ")
For t% = 0 To UBound(ArrayTmp())
    UserData(ID, 17) = Trim(Replace(UserData(ID, 17), ArrayTmp(t%), ""))
    UserData(ID, 17) = UserData(ID, 17) + " " + ArrayTmp(t%)
Next t%
Debug.Print Str(ID) + "[" + UserData(ID, 17) + "]"
End Sub

Sub ItemFlag_Remove(ID, ItemNo)
Dim ArrayTmp() As String
ArrayTmp() = Split(ItemArray(ItemNo, 7), " ")
For t% = 0 To UBound(ArrayTmp())
    UserData(ID, 17) = Trim(Replace(UserData(ID, 17), ArrayTmp(t%), ""))
Next t%
Debug.Print Str(ID) + "[" + UserData(ID, 17) + "]"

End Sub
