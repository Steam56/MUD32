Attribute VB_Name = "Module4"
Global NPCInfo(100, 25) As String
Global NPCCount As Integer

Sub NPC_Load()

Log "Loading NPCs..."

NPCFile$ = Dir(App.Path + "\bots\*.ini")
NPCCount = 0
Do
    NPCFile$ = "bots\" + NPCFile$
    NPCInfo(NPCCount, 0) = ReadINI(NPCFile$, "Name")
    NPCInfo(NPCCount, 1) = ReadINI(NPCFile$, "Description")
    NPCInfo(NPCCount, 2) = ReadINI(NPCFile$, "Health") 'successful attacks deduct items damage value from this.
    NPCInfo(NPCCount, 3) = ReadINI(NPCFile$, "Strength") 'this is how much damage (max) an NPCs attack can do.
    NPCInfo(NPCCount, 4) = ReadINI(NPCFile$, "Location")
    NPCInfo(NPCCount, 5) = UCase$(ReadINI(NPCFile$, "AllowedZone"))
    NPCInfo(NPCCount, 6) = ReadINI(NPCFile$, "ArriveMess")
    NPCInfo(NPCCount, 7) = ReadINI(NPCFile$, "DepartMess")
    NPCInfo(NPCCount, 8) = ReadINI(NPCFile$, "IsHereMess")
    NPCInfo(NPCCount, 9) = ReadINI(NPCFile$, "FullName")
    NPCInfo(NPCCount, 10) = ReadINI(NPCFile$, "MoveSpeed") '0 to 10 (Higher = slower movement, 0 = never move)
    NPCInfo(NPCCount, 11) = ReadINI(NPCFile$, "Value") 'If killed, how much is added to players score
    NPCInfo(NPCCount, 12) = ReadINI(NPCFile$, "Aggression") '0 to 100 (higher = more sucessfull attacks)
    
    NPCInfo(NPCCount, 16) = ReadINI(NPCFile$, "AttackSpeed") '0 to 10 (higher = faster attacks)
    NPCInfo(NPCCount, 17) = ReadINI(NPCFile$, "Respawn") 'In minutes
    NPCInfo(NPCCount, 19) = ReadINI(NPCFile$, "AntiSocial") '0 to 100 (higher = will attack players)
    NPCInfo(NPCCount, 20) = ReadINI(NPCFile$, "AttackMess") 'the message sent to the player when the NPC attacks.
    NPCInfo(NPCCount, 21) = ReadINI(NPCFile$, "NoAttack") '0 = can be attacked, 1 = can't be attacked
    
    NPCInfo(NPCCount, 13) = "0" '0=No-Combat, 1=Combat-Attack, 2=Combat-Victim 3=Dead
    NPCInfo(NPCCount, 14) = Val(NPCInfo(NPCCount, 2)) 'Dynamic Health
    NPCInfo(NPCCount, 15) = "0" 'ID of player it's fighting or interacting with
    NPCInfo(NPCCount, 18) = Val(NPCInfo(NPCCount, 17)) 'Respawn Countdown
    
    NPCFile$ = Dir()
    If NPCFile$ = "" Then Exit Do
    NPCCount = NPCCount + 1
Loop

Log Trim$(Str$(NPCCount + 1)) + " NPCs Loaded."

End Sub

Sub NPC_LocLook(ID, PlayerLoc)

For t% = 0 To NPCCount
    If Val(NPCInfo(t%, 4)) = PlayerLoc Then
        If NPCInfo(t%, 13) = "3" Then
            SendData ID, AnsiTable(Val(UserData(ID, 21)), 9) + "The corpse of " + NPCInfo(t%, 9) + " lies here." + AnsiTable(Val(UserData(ID, 21)), 0)
        ElseIf Val(NPCInfo(t%, 13)) > 0 Then
            SendData ID, AnsiTable(Val(UserData(ID, 21)), 9) + NPCInfo(t%, 9) + " is here, in the middle of a fight!" + AnsiTable(Val(UserData(ID, 21)), 0)
        Else
            SendData ID, AnsiTable(Val(UserData(ID, 21)), 9) + NPCInfo(t%, 8) + AnsiTable(Val(UserData(ID, 21)), 0)
        End If
    End If
Next t%

End Sub


Function NPC_Examine(PlayerLoc, NPCName$)

For t% = 0 To NPCCount
    If Val(NPCInfo(t%, 4)) = PlayerLoc Then
        If UCase$(NPCName$) = UCase$(NPCInfo(t%, 0)) Then
            NPC_Examine = NPCInfo(t%, 1)
            If NPCInfo(t%, 13) = "3" Then NPC_Examine = NPC_Examine + " " + NPCInfo(t%, 9) + " is dead."
            Exit Function
        End If
    End If
Next t%
NPC_Examine = ""

End Function

Sub NPC_TalkTo(ID, NPCName$, Mess$)
For t% = 0 To NPCCount
    If UCase$(NPCName$) = UCase$(NPCInfo(t%, 0)) Then
        Resp$ = NPC_GetResponse(NPCName$, Mess$)
        Resp$ = Replace(Resp$, "%PLAYERNAME%", UserData(ID, 0))
        SendData ID, NPCInfo(t%, 9) + " tells you, " + Chr$(34) + Resp$ + Chr$(34)
    End If
Next t%
End Sub

Function NPC_IsHere(PlayerLoc, NPCName$)
NPC_IsHere = ""
For t% = 0 To NPCCount
    If PlayerLoc = Val(NPCInfo(t%, 4)) Then
        If UCase$(NPCName$) = UCase$(NPCInfo(t%, 0)) Then
            NPC_IsHere = NPCInfo(t%, 0)
            Exit Function
        End If
    End If
Next t%
End Function
Function NPC_ID(NPCName$)
NPC_ID = -1
For t% = 0 To NPCCount
    If UCase$(NPCName$) = UCase$(NPCInfo(t%, 0)) Then
        NPC_ID = t%
        Exit Function
    End If
Next t%
End Function
Function Match(Mess$, PM$)

Dim Elements() As String
Mess$ = UCase$(Mess$)
PM$ = UCase$(PM$)
If PM$ = "NULL" Then Match = 1: Exit Function

Elements = Split(PM$, " ")
For t% = 0 To UBound(Elements)
    If InStr(Mess$, Elements(t%)) <> 0 Then MatchCount = MatchCount + 1
Next t%
If MatchCount = UBound(Elements) + 1 Then Match = 1

End Function
Function NPC_GetResponse(RespFile$, InputMsg$)
Dim Loader(100) As String
Dim LoaderCount As Integer
Dim RespKey() As String

NPC_GetResponse = ""

If IfExist(App.Path + "\bots\mud32.rsp") = True Then
    Handle = FreeFile
    Open App.Path + "\bots\mud32.rsp" For Input Access Read As #Handle
    Do Until EOF(Handle)
        Line Input #Handle, Ln$
        RespKey() = Split(Ln$, "|")
        If RespKey(0) = "*" Then DefResp = RespKey(1)
        If InStr(UCase(InputMsg$), UCase(RespKey(0))) <> 0 Then
            Loader(LoaderCount) = RespKey(1)
            LoaderCount = LoaderCount + 1
        End If
    Loop
    Close #Handle
End If

If IfExist(App.Path + "\bots\" + RespFile$ + ".rsp") = True Then
    Handle = FreeFile
    Open App.Path + "\bots\" + RespFile$ + ".rsp" For Input Access Read As #Handle
    Do Until EOF(Handle)
        Line Input #Handle, Ln$
        RespKey() = Split(Ln$, "|")
        If RespKey(0) = "*" Then DefResp = RespKey(1)
        If InStr(UCase(InputMsg$), UCase(RespKey(0))) <> 0 Then
            Loader(LoaderCount) = RespKey(1)
            LoaderCount = LoaderCount + 1
        End If
    Loop
    Close #Handle
End If

If LoaderCount > 0 Then
    i = Int(Rnd(1) * LoaderCount)
    NPC_GetResponse = Loader(i)
End If

If NPC_GetResponse = "" Then NPC_GetResponse = DefResp

End Function

Sub NPC_Action()
    For t% = 0 To NPCCount
    
        If NPCInfo(t%, 13) = "0" Then 'Its not fighting so...
            'Move it...
            i = Int(Rnd(1) * Val(NPCInfo(t%, 10))) ' Each NPC has own speed, higher speed value = slower the NPC
            If i = 1 Then 'We're going move NPC(T%)
                GetExits 65, Val(NPCInfo(t%, 4))
                j = Int(Rnd(1) * TotalPlayerExit(65))
                If PlayerExit(65, j, 2) > 0 And (CheckLocZone(PlayerExit(65, j, 2)) = NPCInfo(t%, 5)) And (PlayerExit(65, j, 2) <> NPCInfo(t%, 4)) Then
                    uTmp$ = Replace(NPCInfo(t%, 7), "%s", LCase(PlayerExit(65, j, 1)))
                    If InStr(GetLocFlags(NPCInfo(t%, 4)), "DARK") = 0 Then
                        NPC_SysEmoteToRoom t%, NPCInfo(t%, 4), uTmp$ 'Leaves loc A
                    End If
                    NPCInfo(t%, 4) = Str(PlayerExit(65, j, 2))
                    If InStr(GetLocFlags(NPCInfo(t%, 4)), "DARK") = 0 Then
                        NPC_SysEmoteToRoom t%, NPCInfo(t%, 4), NPCInfo(t%, 6) 'Arrives Loc B
                    End If
                End If
            End If
            'Check agression, and attack a player.
            AntiSocial = Int(Rnd(1) * 100)
            If AntiSocial <= Val(NPCInfo(t%, 19)) And Val(NPCInfo(t%, 19)) > 0 Then
                'Pick a player
                Who = -1
                For r% = 0 To 64
                    If Val(UserData(r%, 10)) = Val(NPCInfo(t%, 4)) And Val(UserData(r%, 14)) = 0 And ConnInfo(r%, 1) = "99" Then
                        'player in same room
                        j = Int(Rnd(1) * 2) 'get a one or zero
                        If j = 1 Then
                            Who = r%
                            r% = 65
                        End If
                    End If
                Next r%
                
                'Kick off a fight
                If Who <> -1 Then
                    UserTemp(Who) = "VICTIM" + NPCInfo(t%, 0) + Chr$(0)
                    UserData(Who, 14) = t% + 1
                    If NPCInfo(t%, 13) = "0" Then NPCInfo(t%, 13) = "1"
                    If NPCInfo(t%, 15) = "0" Then NPCInfo(t%, 15) = Trim$(Str$(Who))
                    
                    NPC_FightMsg$ = NPCInfo(t%, 9) + " starts a fight with " + UserData(Who, 0) + " !"
                    Who_FightMsg$ = NPCInfo(t%, 20)
           
                    FightMsgToRoom Val(NPCInfo(t%, 4)), NPC_FightMsg$
                    SendData Who, Who_FightMsg$
                    SayPrompt Who, ""
                End If
            End If
        End If
 
        If NPCInfo(t%, 13) = "1" Or NPCInfo(t%, 13) = "2" Then 'We're Fighting!
            'It's fighting 'who'
            Who = Val(NPCInfo(t%, 15))
            'i = Int(Rnd(1) * 2)
            DoAttack = Int(Rnd(1) * 10)         'Pick a random value 1 to 10
            AttackSpeed = Val(NPCInfo(t%, 16))  'Take Speed

            If (DoAttack < AttackSpeed) Or AttackSpeed = 0 Then Exit Sub      'if Yes/no attack is < speed*4 then quit
        
            Aggressive = Int(Rnd(1) * 100)
            If Aggressive <= Val(NPCInfo(t%, 12)) Then AttackOK = 1 Else AttackOK = 0
           
            If AttackOK = 1 Then 'successful attack
                Damage = Int(Rnd(1) * Val(NPCInfo(t%, 3)))
                AttackMethod = Int(Rnd(1) * 4)
                Select Case AttackMethod
                Case 0
                    AttackMsgA = "delivers a glancing blow"
                Case 1
                    AttackMsgA = "swings low"
                Case 2
                    AttackMsgA = "catches you off guard"
                Case 3
                    AttackMsgA = "parries"
                Case Else
                    AttackMsgA = "ducks and jabs"
                End Select
                
                NPC_FightMsg$ = NPCInfo(t%, 9) + " successfully attacks " + UserData(Who, 0) + " !"
                Who_FightMsg$ = NPCInfo(t%, 9) + " " + AttackMsgA + ", causing " & Damage & " damage!"
            Else
                Damage = 0
                NPC_FightMsg$ = NPCInfo(t%, 9) + " attacks " + UserData(Who, 0) + ", but misses !"
                Who_FightMsg$ = NPCInfo(t%, 9) + " attacks you, but misses!"
            End If
            FightMsgToRoom Val(NPCInfo(t%, 4)), NPC_FightMsg$
            SendData Who, Who_FightMsg$
            
            
            If Damage > 0 Then
                Health = Val(UserData(Who, 11))
                Health = Health - Damage: If Health < 0 Then Health = 0
                UserData(Who, 11) = Str(Health)
                SendData Who, "* Your Health is now " + Trim$(UserData(Who, 11)) + "."
            
                If Health = 0 Then
                  
                    SayFile Who, "misc\death.txt"
                
                    For xx% = 0 To 64
                        If UserData(xx%, 0) <> "" And xx% <> Who Then
                            SendData xx%, "* " + UserData(Who, 0) + " the " + UserData(Who, 4) + PlayerLevel(Who, UserData(Who, 3)) + " has been killed by " + NPCInfo(t%, 9) + "!"
                            SayPrompt xx%, ""
                        End If
                    Next xx%
                
                    SendDDE UserData(Who, 0) + " has been killed by " + NPCInfo(t%, 9) + "!"
                
                
                    ' Sort out dead persons score.
                   
                    OldLevel = PlayerLevel(Who, UserData(Who, 3))
                    Score = Val(UserData(Who, 3))
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
                    NPCInfo(t%, 13) = "0"
                    NPCInfo(t%, 15) = "0"
                    UserData(Who, 14) = 0
                    UserTemp(Who) = ""
                    
                    ReincarnatePlayer Who
                   
                End If
            End If
            SayPrompt Who, ""
         End If
    
    Next t%
End Sub

Sub NPC_SysEmoteToRoom(ID, PlayerLoc, Mess$)
For t% = 0 To 64
    If (Val(UserData(t%, 10)) = PlayerLoc) And ConnInfo(t%, 1) = "99" Then
        Text$ = "* " + Mess$
        SendData t%, Text$
        SayPrompt t%, ""
    End If
Next t%

End Sub

Sub NPC_Attack(ID, Weapon, Damage, NPCID)

    'Setup the NPC and the player so we know whos fighting who
    
    UserTemp(ID) = "ATTACK" + NPCInfo(NPCID, 0) + Chr$(0)
    UserData(ID, 14) = NPCID + 1
    If NPCInfo(NPCID, 13) = "0" Then NPCInfo(NPCID, 13) = "2"
    If NPCInfo(NPCID, 15) = "0" Then NPCInfo(NPCID, 15) = Trim$(Str$(ID))

    If Success(ID, 50, Damage) = True Then
        FightText$ = UserData(ID, 0) + " attacks " + NPCInfo(NPCID, 9) + " with " + Weapon + "!"
        SendData ID, "You attack " + NPCInfo(NPCID, 9) + " with " + Weapon + ", inflicting" + Str$(Damage) + " damage!"
        NPCInfo(NPCID, 14) = NPCInfo(NPCID, 14) - Damage
        If NPCInfo(NPCID, 14) < 0 Then NPCInfo(NPCID, 14) = 0
    Else
        FightText$ = UserData(ID, 0) + " attacks " + NPCInfo(NPCID, 9) + ", but misses !"
        SendData ID, "You attack " + NPCInfo(NPCID, 9) + " with " + Weapon + ", but miss !"
    End If
    
    FightMsgToRoom Val(UserData(ID, 10)), FightText$
    
    If NPCInfo(NPCID, 14) = 0 Then
        SendData ID, "Your last blow to " + NPCInfo(NPCID, 9) + " proves fatal. " + NPCInfo(NPCID, 9) + " collapses to the floor and dies."
        'SendData ID, "Moments later the corpse of " + NPCInfo(NPCID, 9) + " crumbles into dust and blows away on the breeze."
        
        FightMsgToRoom Val(UserData(ID, 10)), UserData(ID, 0) + " has killed " + NPCInfo(NPCID, 9) + "!"
        
        UserTemp(ID) = ""
        UserData(ID, 14) = 0
        NPCInfo(NPCID, 13) = "3"
        NPCInfo(NPCID, 15) = "0"
        
        OldLevel = PlayerLevel(ID, UserData(ID, 3))
        Score = Val(UserData(ID, 3))
        Score = Score + Val(NPCInfo(NPCID, 11))
        UserData(ID, 3) = Trim$(Str$(Score))
        SendData ID, "Your score is now " + Format$(Val(UserData(ID, 3)), "###,###,##0") + "."
        NewLevel = PlayerLevel(ID, UserData(ID, 3))
        Call UpdateTopTen(UserData(ID, 0), UserData(ID, 3), UserData(ID, 2))
        If OldLevel <> NewLevel Then
            Text$ = "Congratulations, you are now " + UserData(ID, 0) + " the " + UserData(ID, 4) + PlayerLevel(ID, Val(UserData(ID, 3))) + "!"
            SendData ID, Text$
        End If
    End If
    
        
End Sub

Sub NPC_Respawn()

For t% = 0 To NPCCount
        If NPCInfo(t%, 13) = "3" Then
            NPCInfo(t%, 18) = NPCInfo(t%, 18) - 1
            If NPCInfo(t%, 18) = 0 Then  'respawn the dead npc
                NPCInfo(t%, 18) = Val(NPCInfo(t%, 17))
                NPCInfo(t%, 13) = "0"
                NPCInfo(NPCCount, 14) = Val(NPCInfo(NPCCount, 2))

                For r% = 0 To 64
                    If Val(UserData(r%, 10)) = Val(NPCInfo(t%, 4)) Then
                        Text$ = "* The corpse of " + NPCInfo(t%, 9) + " starts to shimmer... " + vbCrLf + "* " + NPCInfo(t%, 9) + " has been resurrected."
                        SendData r%, Text$
                        SayPrompt r%, ""
                    End If
                Next r%
            End If
        End If

Next t%

End Sub
