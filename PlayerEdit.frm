VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUD32 - PlayerEdit"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "PlayerEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   6930
      TabIndex        =   106
      Top             =   7425
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   50
      Left            =   5895
      TabIndex        =   104
      Text            =   "Text1"
      Top             =   6975
      Width           =   2220
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Change"
      Height          =   285
      Left            =   3060
      TabIndex        =   103
      Top             =   765
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Load"
      Height          =   285
      Left            =   3060
      TabIndex        =   101
      Top             =   225
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   330
      Left            =   5625
      TabIndex        =   100
      Top             =   7425
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   49
      Left            =   5895
      TabIndex        =   99
      Text            =   "Text1"
      Top             =   6705
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   48
      Left            =   5895
      TabIndex        =   98
      Text            =   "Text1"
      Top             =   6435
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   47
      Left            =   5895
      TabIndex        =   97
      Text            =   "Text1"
      Top             =   6165
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   46
      Left            =   5895
      TabIndex        =   96
      Text            =   "Text1"
      Top             =   5895
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   45
      Left            =   5895
      TabIndex        =   95
      Text            =   "Text1"
      Top             =   5625
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   44
      Left            =   5895
      TabIndex        =   94
      Text            =   "Text1"
      Top             =   5355
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   43
      Left            =   5895
      TabIndex        =   93
      Text            =   "Text1"
      Top             =   5085
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   42
      Left            =   5895
      TabIndex        =   92
      Text            =   "Text1"
      Top             =   4815
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   41
      Left            =   5895
      TabIndex        =   91
      Text            =   "Text1"
      Top             =   4545
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   40
      Left            =   5895
      TabIndex        =   90
      Text            =   "Text1"
      Top             =   4275
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   39
      Left            =   5895
      TabIndex        =   89
      Text            =   "Text1"
      Top             =   4005
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   38
      Left            =   5895
      TabIndex        =   88
      Text            =   "Text1"
      Top             =   3735
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   37
      Left            =   5895
      TabIndex        =   87
      Text            =   "Text1"
      Top             =   3465
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   36
      Left            =   5895
      TabIndex        =   86
      Text            =   "Text1"
      Top             =   3195
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   35
      Left            =   5895
      TabIndex        =   85
      Text            =   "Text1"
      Top             =   2925
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   34
      Left            =   5895
      TabIndex        =   84
      Text            =   "Text1"
      Top             =   2655
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   33
      Left            =   5895
      TabIndex        =   83
      Text            =   "Text1"
      Top             =   2385
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   32
      Left            =   5895
      TabIndex        =   82
      Text            =   "Text1"
      Top             =   2115
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   31
      Left            =   5895
      TabIndex        =   81
      Text            =   "Text1"
      Top             =   1845
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   30
      Left            =   5895
      TabIndex        =   80
      Text            =   "Text1"
      Top             =   1575
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   29
      Left            =   5895
      TabIndex        =   79
      Text            =   "Text1"
      Top             =   1305
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   28
      Left            =   5895
      TabIndex        =   78
      Text            =   "Text1"
      Top             =   1035
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   27
      Left            =   5895
      TabIndex        =   77
      Text            =   "Text1"
      Top             =   765
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   5895
      TabIndex        =   76
      Text            =   "Text1"
      Top             =   495
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   5895
      TabIndex        =   75
      Text            =   "Text1"
      Top             =   225
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   1800
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   6975
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   1800
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   6705
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   1800
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   6435
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   21
      Left            =   1800
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   6165
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   1800
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   5895
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   1800
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   5625
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   1800
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   5355
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   1800
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   5085
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   1800
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   4815
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   15
      Left            =   1800
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   4545
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   14
      Left            =   1800
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   4275
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   1800
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   4005
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   1800
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   3735
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   1800
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   3465
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   1800
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   3195
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   2925
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1800
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   2655
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1800
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   2385
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   2115
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   1845
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1575
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   1305
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1035
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   765
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   225
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   50
      Left            =   4230
      TabIndex        =   105
      Top             =   7020
      Width           =   1590
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   240
      Left            =   180
      TabIndex        =   102
      Top             =   7560
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   49
      Left            =   4230
      TabIndex        =   74
      Top             =   6750
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   48
      Left            =   4230
      TabIndex        =   73
      Top             =   6480
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   47
      Left            =   4230
      TabIndex        =   72
      Top             =   6210
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   46
      Left            =   4230
      TabIndex        =   71
      Top             =   5940
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   45
      Left            =   4230
      TabIndex        =   70
      Top             =   5670
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   44
      Left            =   4230
      TabIndex        =   69
      Top             =   5400
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   43
      Left            =   4230
      TabIndex        =   68
      Top             =   5130
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   42
      Left            =   4230
      TabIndex        =   67
      Top             =   4860
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   41
      Left            =   4230
      TabIndex        =   66
      Top             =   4590
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   40
      Left            =   4230
      TabIndex        =   65
      Top             =   4320
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   39
      Left            =   4230
      TabIndex        =   64
      Top             =   4050
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   38
      Left            =   4230
      TabIndex        =   63
      Top             =   3780
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   37
      Left            =   4230
      TabIndex        =   62
      Top             =   3510
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   36
      Left            =   4230
      TabIndex        =   61
      Top             =   3240
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   35
      Left            =   4230
      TabIndex        =   60
      Top             =   2970
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   34
      Left            =   4230
      TabIndex        =   59
      Top             =   2700
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   33
      Left            =   4230
      TabIndex        =   58
      Top             =   2430
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   32
      Left            =   4230
      TabIndex        =   57
      Top             =   2160
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   31
      Left            =   4230
      TabIndex        =   56
      Top             =   1890
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   30
      Left            =   4230
      TabIndex        =   55
      Top             =   1620
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   29
      Left            =   4230
      TabIndex        =   54
      Top             =   1350
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   28
      Left            =   4230
      TabIndex        =   53
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   27
      Left            =   4230
      TabIndex        =   52
      Top             =   810
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   26
      Left            =   4230
      TabIndex        =   51
      Top             =   540
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   25
      Left            =   4230
      TabIndex        =   50
      Top             =   270
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   24
      Left            =   135
      TabIndex        =   24
      Top             =   7020
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   23
      Left            =   135
      TabIndex        =   23
      Top             =   6750
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   22
      Left            =   135
      TabIndex        =   22
      Top             =   6480
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   21
      Left            =   135
      TabIndex        =   21
      Top             =   6210
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   20
      Left            =   135
      TabIndex        =   20
      Top             =   5940
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   19
      Left            =   135
      TabIndex        =   19
      Top             =   5670
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   18
      Left            =   135
      TabIndex        =   18
      Top             =   5400
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   17
      Left            =   135
      TabIndex        =   17
      Top             =   5130
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   16
      Left            =   135
      TabIndex        =   16
      Top             =   4860
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   15
      Left            =   135
      TabIndex        =   15
      Top             =   4590
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   14
      Left            =   135
      TabIndex        =   14
      Top             =   4320
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   13
      Left            =   135
      TabIndex        =   13
      Top             =   4050
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   12
      Left            =   135
      TabIndex        =   12
      Top             =   3780
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   11
      Left            =   135
      TabIndex        =   11
      Top             =   3510
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   10
      Top             =   3240
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   9
      Left            =   135
      TabIndex        =   9
      Top             =   2970
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   8
      Top             =   2700
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   7
      Top             =   2430
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   6
      Top             =   2160
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   5
      Top             =   1890
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   4
      Top             =   1620
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   1350
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   1
      Top             =   810
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   270
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveUser Trim$(Form1.Text1(0).Text)
x = MsgBox("Saved!")

End Sub

Private Sub Command2_Click()
LoadUser Form1.Text1(0).Text
End Sub

Private Sub Command3_Click()
Form1.Label2.Caption = Crypt(Form1.Text1(1).Text)
Form1.Text1(1).Text = String(Len(Form1.Text1(1).Text), "*")
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
For T% = 0 To 50
    Form1.Label1(T%).Caption = Trim$(Str$(T%)) + "."
    Form1.Label1(T%).Alignment = 1
    Form1.Text1(T%).Text = ""
Next T%
Form1.Label1(0).Caption = "Player 0."
Form1.Label1(1).Caption = "Password 1."
Form1.Label1(2).Caption = "Gender 2."
Form1.Label1(3).Caption = "Points 3."
Form1.Label1(4).Caption = "Suffix 4."
Form1.Label1(5).Caption = "Brief 5."
Form1.Label1(6).Caption = "Profile 6."
Form1.Label1(7).Caption = "Echo 7."
Form1.Label1(8).Caption = "StatusBar 8."
Form1.Label1(9).Caption = "Builder/Immort 9."
Form1.Label1(10).Caption = "Location 10."
Form1.Label1(11).Caption = "Health 11."
Form1.Label1(15).Caption = "Strength 15."
Form1.Label1(21).Caption = "Colour 21."
Form1.Label1(22).Caption = "Term Width 22."
End Sub

Sub LoadUser(Username$)
Handle = FreeFile
x = Dir(App.Path + "\players\" + Username$ + ".plr")
If x = "" Then
    Beep
    Exit Sub
End If
Open App.Path + "\players\" + Username$ + ".plr" For Input Access Read As #Handle
    Line Input #Handle, Ln$
    Form1.Text1(0).Text = Ln$
    Line Input #Handle, Ln$
    Form1.Label2.Caption = Ln$
    Form1.Text1(1).Text = String(Len(Ln$), "*")
    For T% = 2 To 50
        Line Input #Handle, Ln$
        Form1.Text1(T%).Text = Ln$
    Next T%
Close Handle
End Sub

Function Crypt(Pass$)
For T% = 1 To Len(Pass$)
    x% = Asc(Mid$(Pass$, T%, 1))
    Y% = x% Xor (27 + T%)
    NewPass$ = NewPass$ + Chr$(Y%)
Next T%
Crypt = NewPass$

End Function

Sub SaveUser(Username$)
Handle = FreeFile
Open App.Path + "\players\" + Username$ + ".plr" For Output Access Write As #Handle
    Print #Handle, Form1.Text1(0).Text
    Print #Handle, Form1.Label2.Caption
    For T% = 2 To 50
        Print #Handle, Form1.Text1(T%).Text
    Next T%
Close Handle
End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub
