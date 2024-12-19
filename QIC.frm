VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUD32 - Quick Item Creator"
   ClientHeight    =   6255
   ClientLeft      =   1185
   ClientTop       =   855
   ClientWidth     =   7365
   Icon            =   "QIC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7365
   Begin VB.CommandButton Command5 
      Caption         =   "Find"
      Height          =   315
      Left            =   2220
      TabIndex        =   31
      Top             =   5760
      Width           =   555
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   3000
      TabIndex        =   28
      Top             =   4680
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   25
      Top             =   4080
      Width           =   4215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&New"
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Erase"
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   5760
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   360
      TabIndex        =   19
      Top             =   6480
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5280
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   6480
      TabIndex        =   7
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   4140
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   5280
      TabIndex        =   6
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   4
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Top             =   2880
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Index           =   2
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2040
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   765
      Index           =   1
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Flags (LIGHT POISON INVIS etc)"
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   29
      Top             =   4440
      Width           =   4095
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   6840
      TabIndex        =   27
      Top             =   5325
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "No."
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   26
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Too Heavy Message (optional)"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   24
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Starts in Location"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   18
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Respawn"
      Height          =   255
      Index           =   3
      Left            =   6480
      TabIndex        =   17
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Damage"
      Height          =   255
      Index           =   2
      Left            =   4140
      TabIndex        =   16
      Top             =   3240
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Value"
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Weight"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   14
      Top             =   3240
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Short Names (Comma separated)"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   13
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Location Description"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   12
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Description (Exam)"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   11
      Top             =   720
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Full Name"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim locarray$(1000)
Dim locval(1000)
Dim itemarray$(1000)
Dim itemval(1000)


Private Sub Combo1_Click()

For T% = 0 To 1000
If locarray$(T%) = Form1.Combo1.Text Then
        Form1.Label3.Caption = locval(T%)
End If
Next T%

End Sub

Private Sub Command1_Click()
SaveItem
End Sub

Private Sub Command2_Click()
xName$ = Form1.Text1(0).Text
For T% = 0 To 1000
    If itemarray$(T%) = xName$ Then
        deadfile$ = App.Path + "\items\" + Trim$(Str$(itemval(T%))) + ".itx"
        Kill deadfile$
        ItemLoad
        For r% = 0 To 7
            Form1.Text1(r%).Text = ""
            Form1.Text1(7).Text = "20"
            Form1.Text1(6).Text = "0"
        Next r%
        T% = 1001
    End If
Next T%
        
    

End Sub

Private Sub Command3_Click()
For T% = 0 To 7
Form1.Text1(T%).Text = ""
Form1.Text1(7).Text = "20"
Form1.Text1(6).Text = "0"
Next T%
End Sub

Private Sub Command4_Click()
End

End Sub

Private Sub Command5_Click()
X = Form1.List1.ListIndex
If X > 0 Then X = X + 1
If X > Form1.List1.ListCount Then X = Form1.List1.ListCount
For T% = X To Form1.List1.ListCount
    If InStr(UCase$(Form1.List1.List(T%)), UCase$(Form1.Text2.Text)) <> 0 Then
        Form1.List1.ListIndex = T%
        T% = Form1.List1.ListCount + 1
    End If
Next T%

End Sub

Private Sub Form_Load()
Form2.Top = (Screen.Height - Form2.Height) / 2
Form2.Left = (Screen.Width - Form2.Width) / 2
Form2.Show
DoEvents
LocLoad

ItemLoad






Form2.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub List1_Click()
X$ = Form1.List1.List(Form1.List1.ListIndex)
For T% = 0 To 1000
If itemarray$(T%) = X$ Then Found = itemval(T%)
Next T%
Filename$ = "items\" + (Trim$(Str$(Found))) + ".itx"
Form1.Text1(0).Text = ReadINI(Filename$, "fullname")
Form1.Text1(1).Text = ReadINI(Filename$, "desc")
Form1.Text1(2).Text = ReadINI(Filename$, "locdesc")
Form1.Text1(3).Text = ReadINI(Filename$, "shortname")
Form1.Text1(4).Text = ReadINI(Filename$, "weight")
Form1.Text1(5).Text = ReadINI(Filename$, "value")
Form1.Text1(6).Text = ReadINI(Filename$, "damage")
Form1.Text1(7).Text = ReadINI(Filename$, "respawn")
Form1.Text1(8).Text = ReadINI(Filename$, "tooheavy")
Form1.Text1(9).Text = ReadINI(Filename$, "flags")

Y% = Val(ReadINI(Filename$, "location"))
Found = 0
For T% = 0 To 1000
If locval(T%) = Y% Then
    If locarray$(T%) <> "" Then
        Found = 1
        Form1.Combo1.Text = locarray$(T%)
        Form1.Label3.Caption = locval(T%)
        T% = 1001
    End If
End If

Next T%
If Found <> 1 Then
    Form1.Combo1.Text = "<Undefined Location>"
    Form1.Label3.Caption = Str$(Y%)
End If

End Sub

Function ReadINI(Filename$, Keyword$)

' Called by: Value = ReadINI(INIFile, Keyword)

Handle = FreeFile
Open App.Path + "\" + Filename$ For Input Access Read As #Handle
Do Until EOF(Handle)
Line Input #Handle, ln$
ln$ = Trim$(ln$)
If ln$ <> "" And Left$(ln$, 1) <> ";" Then
    If UCase$(Left$(ln$, Len(Keyword$))) = UCase$(Keyword$) Then
        I% = InStr(ln$, "=")
        Tmp$ = Mid$(ln$, I% + 1)
        ReadINI = Trim$(Tmp$)
        Close Handle
        Exit Function
    End If
End If
Loop
ReadINI = ""
Close Handle

End Function
Sub ItemLoad()
For T% = 0 To 1000
itemarray$(T%) = ""
itemval(T%) = ""
Next T%
Form1.List1.Clear
Form1.List1.Refresh
'load items
Form1.File1.Path = App.Path + "\items"
Form1.File1.Pattern = "*.itx"
Form1.File1.Refresh
For T% = 0 To Form1.File1.ListCount - 1
Filename$ = Form1.File1.List(T%)
Locit = Val(Left$(Filename$, Len(Filename$) - 4))
Open App.Path + "\items\" + Filename$ For Input Access Read As #1
Do Until EOF(1)
Line Input #1, ln$
If UCase$(Left$(ln$, 8)) = "FULLNAME" Then
    ln$ = Trim$(Mid$(ln$, InStr(ln$, "=") + 1))
    itemarray$(T%) = ln$
    Form1.List1.AddItem ln$
    Form2.Label4.Caption = ln$
    Form2.Label4.Refresh
    itemval(T%) = Locit
End If
Loop
Close #1
Next T%
End Sub
Sub LocLoad()
'load Locations
Form1.Combo1.AddItem "<Undefined Location>"
Form1.File1.Path = App.Path + "\locations"
Form1.File1.Pattern = "*.loc"
For T% = 0 To Form1.File1.ListCount - 1
Filename$ = Form1.File1.List(T%)
Locit = Val(Left$(Filename$, Len(Filename$) - 4))
Open App.Path + "\locations\" + Filename$ For Input Access Read As #1
Do Until EOF(1)
Line Input #1, ln$
If Left$(ln$, 1) = "0" Then
    ln$ = Mid$(ln$, 3)
    locarray$(T%) = ln$
    Form1.Combo1.AddItem ln$
    Form2.Label4.Caption = ln$
    Form2.Label4.Refresh
    locval(T%) = Locit
End If
Loop
Close #1
Next T%
End Sub

Sub SaveItem()
xName$ = Form1.Text1(0).Text
For T% = 0 To 1000
    If itemarray$(T%) = xName$ Then
        overfile$ = App.Path + "\items\" + Trim$(Str$(itemval(T%))) + ".itx"
        T% = 1001
    End If
Next T%
If overfile$ = "" Then
    For T% = 0 To 1000
    freeslot = T%
    Taken = 0
    For r% = 0 To 50
        If itemval(r%) = freeslot Then
            Taken = 1
        End If
    Next r%
    If Taken <> 1 Then T% = 1001
    Next T%
    overfile$ = App.Path + "\items\" + Trim$(Str$(freeslot)) + ".itx"
End If
Open overfile$ For Output Access Write As #1
Print #1, "fullname= " + Form1.Text1(0).Text
Print #1, "desc= " + Form1.Text1(1).Text
Print #1, "locdesc= " + Form1.Text1(2).Text
Print #1, "shortname= " + Form1.Text1(3).Text
Print #1, "weight= " + Form1.Text1(4).Text
Print #1, "value= " + Form1.Text1(5).Text
Print #1, "damage= " + Form1.Text1(6).Text
Print #1, "respawn= " + Form1.Text1(7).Text
If Form1.Text1(8).Text <> "" Then Print #1, "tooheavy= " + Form1.Text1(8).Text
If Form1.Text1(9).Text <> "" Then Print #1, "flags= " + Form1.Text1(9).Text

For T% = 0 To 1000
If locarray$(T%) = Form1.Combo1.Text Then
    Print #1, "location=" + Trim$(Str$(locval(T%)))
End If
Next T%

Close #1

ItemLoad



    

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command5_Click
    KeyAscii = 0
End If

End Sub
