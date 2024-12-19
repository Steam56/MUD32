VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUD32 - NPC Editor"
   ClientHeight    =   6255
   ClientLeft      =   1665
   ClientTop       =   1140
   ClientWidth     =   10515
   Icon            =   "NPCEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Global Dialog"
      Height          =   315
      Index           =   5
      Left            =   6840
      TabIndex        =   46
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NPC &Dialog"
      Height          =   315
      Index           =   4
      Left            =   5520
      TabIndex        =   45
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Details"
      Height          =   5415
      Left            =   1500
      TabIndex        =   25
      Top             =   120
      Width           =   6555
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   38
         Top             =   2340
         Width           =   6075
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   37
         Top             =   3000
         Width           =   6075
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   36
         Top             =   3660
         Width           =   6075
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   35
         Top             =   4320
         Width           =   6075
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   4920
         Width           =   4035
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   4620
         TabIndex        =   33
         Top             =   4920
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   ">"
         Height          =   285
         Left            =   4320
         TabIndex        =   32
         Top             =   4920
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   2235
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   2580
         TabIndex        =   27
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         Height          =   825
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1140
         Width           =   6075
      End
      Begin VB.Label Label1 
         Caption         =   "Arrival Message"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   44
         Top             =   2100
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Departure Message (use %s to include direction)"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   43
         Top             =   2760
         Width           =   3555
      End
      Begin VB.Label Label1 
         Caption         =   "Location Description"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   42
         Top             =   3420
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Attack Message"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   41
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Spawns in Location"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   40
         Top             =   4680
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Allowed Zones"
         Height          =   195
         Index           =   16
         Left            =   4620
         TabIndex        =   39
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Short Name (one word only)"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Full Name"
         Height          =   195
         Index           =   1
         Left            =   2580
         TabIndex        =   30
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label1 
         Caption         =   "Description (Exam)"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   900
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stats"
      Height          =   5415
      Left            =   8100
      TabIndex        =   7
      Top             =   120
      Width           =   2355
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   1560
         TabIndex        =   47
         Top             =   3600
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   1560
         TabIndex        =   22
         Top             =   2880
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1560
         TabIndex        =   21
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Auto Generate Stats"
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1875
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1560
         TabIndex        =   16
         Top             =   1800
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   1560
         TabIndex        =   15
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   1560
         TabIndex        =   14
         Top             =   2520
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   1560
         TabIndex        =   9
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   1560
         TabIndex        =   8
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "No Attack"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   48
         Top             =   3660
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Aggression"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   24
         Top             =   3300
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "AntiSocial"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   23
         Top             =   2940
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Move Speed"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   19
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Attack Speed"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   18
         Top             =   2580
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Respawn"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   17
         Top             =   1860
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Max Health"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   13
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Strength"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   12
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Value"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   11
         Top             =   1500
         Width           =   615
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Quit"
      Height          =   315
      Index           =   3
      Left            =   9120
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Erase"
      Height          =   315
      Index           =   2
      Left            =   4140
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Save"
      Height          =   315
      Index           =   1
      Left            =   2820
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&New"
      Height          =   315
      Index           =   0
      Left            =   1500
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   5445
      IntegralHeight  =   0   'False
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Welcome to The NPC Editor!"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   6000
      Width           =   10395
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NPC Count"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Change()
Form1.Command3(1).Enabled = True

End Sub

Private Sub Command1_Click()
GenerateStats
Form1.Command3(1).Enabled = True

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowHelp 17
End Sub

Private Sub Command2_Click()
PushZone
Form1.Text1(15).SetFocus
End Sub

Private Sub Command3_Click(Index As Integer)
If Index = 0 Then NewNPC
If Index = 1 Then
    SaveNPC
    Form1.Command3(1).Enabled = False
    Form1.Text1(0).SetFocus
End If

If Index = 2 Then
    EraseNPC
End If
If Index = 3 Then
    Unload Form1
End If

If Index = 4 Then If Form3.Visible = True Then Form3.Visible = False Else Form3.Visible = True
If Index = 5 Then If Form4.Visible = True Then Form4.Visible = False Else Form4.Visible = True

End Sub

Private Sub Form_Load()
Form2.Show
LoadNPCList
LoadLocs
Form2.Hide
Form1.List1.ListIndex = 0
Form1.Command3(1).Enabled = False
LoadGBL

End Sub
Sub NewNPC()
For T% = 0 To 15
    Form1.Text1(T%).Text = ""
Next T%
Form1.Text1(0).SetFocus
Form1.List1.ListIndex = -1
Form1.Combo1.ListIndex = 0

End Sub
Sub EraseNPC()
X = MsgBox("Delete NPC named '" + Form1.Text1(1).Text + "' ?", vbYesNo + vbQuestion, App.Title)
If X = vbYes Then
    Kill App.Path + "\bots\" + Form1.Text1(0).Text + ".ini"
    If Dir(App.Path + "\bots\" + Form1.Text1(0).Text + ".rsp") <> "" Then
        Kill App.Path + "\bots\" + Form1.Text1(0).Text + ".rsp"
    End If
    LoadNPCList
    NewNPC
End If

End Sub

Sub LoadNPCList()
Form1.List1.Clear
X = Dir(App.Path + "\bots\*.ini", vbNormal)
Do While X <> ""
    Form1.List1.AddItem Left$(X, Len(X) - 4)
    X = Dir()
Loop
Form1.Label2.Caption = Form1.List1.ListCount & " NPCs"

End Sub

Sub LoadLocs()
Form1.Combo1.Clear
Form1.Combo1.AddItem "<Undefined Location>"
X = Dir(App.Path + "\locations\*.loc", vbNormal)
Do While X <> ""
    Open App.Path + "\locations\" + X For Input Access Read As #1
    Line Input #1, ln$
    If Left$(ln$, 1) = "0" Then
        Form1.Combo1.AddItem Mid$(ln$, 3) + " [" + Left$(X, Len(X) - 4) + "]"
        Form2.Label4.Caption = Mid$(ln$, 3)
        Form2.Label4.Refresh
    End If
    Close #1
    
    X = Dir()
Loop

End Sub
Sub LoadNPC(NPCFile$)
For T% = 0 To 15
    Form1.Text1(T%).Text = ""
Next T%

Form1.Text1(0).Text = ReadINI(NPCFile$, "Name")
Form1.Text1(1).Text = ReadINI(NPCFile$, "FullName")
Form1.Text1(2).Text = ReadINI(NPCFile$, "Description")
Form1.Text1(3).Text = ReadINI(NPCFile$, "ArriveMess")
Form1.Text1(4).Text = ReadINI(NPCFile$, "DepartMess")
Form1.Text1(5).Text = ReadINI(NPCFile$, "IsHereMess")
Form1.Text1(6).Text = ReadINI(NPCFile$, "AttackMess")
Form1.Text1(7).Text = ReadINI(NPCFile$, "Health")
Form1.Text1(8).Text = ReadINI(NPCFile$, "Strength")
Form1.Text1(9).Text = ReadINI(NPCFile$, "Value")
Form1.Text1(10).Text = ReadINI(NPCFile$, "Respawn")
Form1.Text1(11).Text = ReadINI(NPCFile$, "MoveSpeed")
Form1.Text1(12).Text = ReadINI(NPCFile$, "AttackSpeed")
Form1.Text1(13).Text = ReadINI(NPCFile$, "AntiSocial")
Form1.Text1(14).Text = ReadINI(NPCFile$, "Aggression")
Form1.Text1(15).Text = ReadINI(NPCFile$, "AllowedZone")
Form1.Text1(16).Text = ReadINI(NPCFile$, "NoAttack")

LocID = ReadINI(NPCFile, "Location")
Found = 0
For T% = 0 To Form1.Combo1.ListCount - 1
    ListLoc = Form1.Combo1.List(T%)
    ListLoc = Replace(Mid$(ListLoc, InStrRev(ListLoc, " ")), "[", "")
    ListLoc = Replace(Mid$(ListLoc, InStrRev(ListLoc, " ")), "]", "")
    
    If Val(ListLoc) = Val(LocID) Then Found = 1: ListLocID = T%: T% = Form1.Combo1.ListCount + 1

Next T%
If Found = 1 Then
    Form1.Combo1.ListIndex = ListLocID
Else
    Form1.Combo1.ListIndex = 0
End If
Form1.Command3(1).Enabled = False

NPCRespFile = App.Path + "\" + Replace(NPCFile, ".ini", ".rsp")
X = Dir(NPCRespFile)
If X <> "" Then
    Handle = FreeFile
    Open NPCRespFile For Binary Access Read As #Handle
    xBuffer$ = Space(LOF(Handle))
    Get #Handle, , xBuffer$
    Form3.Text1.Text = xBuffer$
    Close Handle
Else
    Form3.Text1.Text = ""
End If
Form3.Caption = "Personal Dialog for: " + Form1.List1.List(Form1.List1.ListIndex)

End Sub
Sub LoadGBL()
GBLRespFile = App.Path + "\bots\mud32.rsp"
X = Dir(GBLRespFile)
If X <> "" Then
    Handle = FreeFile
    Open GBLRespFile For Binary Access Read As #Handle
    xBuffer$ = Space(LOF(Handle))
    Get #Handle, , xBuffer$
    Form4.Text1.Text = xBuffer$
    Close Handle
Else
    Form4.Text1.Text = ""
End If
Form4.Caption = "Global Dialog, All NPCs use this first."

End Sub
Sub SaveGBL()
If Dir(App.Path + "\bots\mud32.rsp") <> "" Then
    Kill App.Path + "\bots\mud32.rsp"
End If
If Trim(Form4.Text1.Text) <> "" Then
    Handle = FreeFile
    Open App.Path + "\bots\mud32.rsp" For Binary Access Write As Handle
    Put #Handle, , Form4.Text1.Text
    Close Handle
End If
End Sub

Function ReadINI(Filename$, Keyword$)

' Called by: Value = ReadINI(INIFile, Keyword)
' this is not good for massive INI files, but will suffice.
' Yes I know about ReadPrivateProfile etc, but thats not very portable is it?

Handle = FreeFile
Open App.Path + "\" + Filename$ For Input Access Read As #Handle
Do Until EOF(Handle)
Line Input #Handle, ln$
ln$ = Trim$(ln$)
If ln$ <> "" And Left$(ln$, 1) <> ";" And Left$(ln$, 1) <> "[" Then
    If UCase$(Left$(ln$, Len(Keyword$))) = UCase$(Keyword$) Then
        i% = InStr(ln$, "=")
        tmp$ = Mid$(ln$, i% + 1)
        ReadINI = Trim$(tmp$)
        Close Handle
        Exit Function
    End If
End If
Loop
ReadINI = ""
Close Handle

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowHelp 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveGBL
Unload Form2
Unload Form3
Unload Form4

End
End Sub

Private Sub List1_Click()
If Form1.List1.ListIndex <> -1 Then LoadNPC "bots\" + Form1.List1.List(Form1.List1.ListIndex) + ".ini"

End Sub

Sub PushZone()
    ListLoc = Form1.Combo1.List(Form1.Combo1.ListIndex)
    ListLoc = Replace(Mid$(ListLoc, InStrRev(ListLoc, " ")), "[", "")
    ListLoc = Trim$(Replace(ListLoc, "]", "") + ".loc")
    Open App.Path + "\locations\" + ListLoc For Input Access Read As #2
    Do Until EOF(2)
        Line Input #2, ln$
        If Left$(ln$, 1) = "4" Then Form1.Text1(15).Text = Mid$(ln$, 3)
    Loop
    Close #2
    

End Sub


Sub SaveNPC()
Open App.Path + "\bots\" + Form1.Text1(0).Text + ".ini" For Output Access Write As #2


Print #2, "Name=" + Form1.Text1(0).Text
Print #2, "FullName=" + Form1.Text1(1).Text
Print #2, "Description=" + Replace(Form1.Text1(2).Text, vbCrLf, " ")
Print #2, "ArriveMess=" + Form1.Text1(3).Text
Print #2, "DepartMess=" + Form1.Text1(4).Text
Print #2, "IsHereMess=" + Form1.Text1(5).Text
Print #2, "AttackMess=" + Form1.Text1(6).Text
Print #2, "Health=" + Form1.Text1(7).Text
Print #2, "Strength=" + Form1.Text1(8).Text
Print #2, "Value=" + Form1.Text1(9).Text
Print #2, "Respawn=" + Form1.Text1(10).Text
Print #2, "MoveSpeed=" + Form1.Text1(11).Text
Print #2, "AttackSpeed=" + Form1.Text1(12).Text
Print #2, "AntiSocial=" + Form1.Text1(13).Text
Print #2, "Aggression=" + Form1.Text1(14).Text
Print #2, "AllowedZone=" + Form1.Text1(15).Text
Print #2, "NoAttack=" + Form1.Text1(16).Text


ListLoc = Form1.Combo1.List(Form1.Combo1.ListIndex)
ListLoc = Replace(Mid$(ListLoc, InStrRev(ListLoc, " ")), "[", "")
ListLoc = Replace(Mid$(ListLoc, InStrRev(ListLoc, " ")), "]", "")

Print #2, "Location=" + Trim(ListLoc)

Close #2
If Dir(App.Path + "\bots\" + Form1.Text1(0).Text + ".rsp") <> "" Then
    Kill App.Path + "\bots\" + Form1.Text1(0).Text + ".rsp"
End If
If Trim(Form3.Text1.Text) <> "" Then
    Handle = FreeFile
    Open App.Path + "\bots\" + Form1.Text1(0).Text + ".rsp" For Binary Access Write As Handle
    Put #Handle, , Form3.Text1.Text
    Close Handle
End If

LoadNPCList
For T% = 0 To Form1.List1.ListCount - 1
    If Form1.List1.List(T%) = Form1.Text1(0).Text Then Form1.List1.ListIndex = T%
Next T%


End Sub

Sub ShowHelp(ID)
Dim HelpMess(20) As String

HelpMess(0) = "Short one word name of the NPC ie. 'woodsman'"
HelpMess(1) = "Full name of the NPC ie. 'The Old Woodsman'"

HelpMess(2) = "Message for 'EXAM <npc>'"
HelpMess(3) = "Message displayed when the NPC enters a room."
HelpMess(4) = "Message displayed when the NPC leaves a room."
HelpMess(5) = "Message displayed when the NPC is in room that the player enters."
HelpMess(6) = "Message sent to the player when the NPC attacks."

HelpMess(7) = "Successful attacks deduct damage value from this, relative to an items 'damage' attrib."
HelpMess(8) = "Maximum damage an NPC's attack can do to a player."
HelpMess(9) = "If NPC is killed, how much is added to players score."
HelpMess(10) = "Time in minutes, before the NPC is re-incarnated."

HelpMess(11) = "Range is 0 to 10 (Higher = slower movement, 0 = never move)"
HelpMess(12) = "Range is 0 to 10 (higher = faster attacks)"
HelpMess(13) = "Range is 0 to 100 (higher = will attack players, 0 = never attack)"
HelpMess(14) = "Range is 0 to 100 (higher = more successful attacks)"

HelpMess(15) = "Set the Zone(s) the NPC is allowed in."
HelpMess(16) = "1 = NPC cannot be attacked, 0 = NPC can be attacked"

HelpMess(17) = "Click to auto-generate random values"
    
Form1.Label3.Caption = HelpMess(ID)

End Sub

Private Sub Text1_Change(Index As Integer)
Form1.Command3(1).Enabled = True

End Sub

Private Sub Text1_GotFocus(Index As Integer)
ShowHelp Index
End Sub

Sub GenerateStats()
For T% = 7 To 14
    Form1.Text1(T%).Text = RndVal(100)
Next T%
For T% = 11 To 12
    Form1.Text1(T%).Text = RndVal(10)
Next T%
Form1.Text1(9).Text = RndVal(100) * 10
Form1.Text1(16).Text = 0


End Sub
Function RndVal(Max)
RndVal = (Int(Rnd(1) * (Max - 1))) + 1
End Function

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowHelp Index

End Sub
