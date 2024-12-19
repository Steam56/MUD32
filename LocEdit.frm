VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUD32 - Location Editor"
   ClientHeight    =   7725
   ClientLeft      =   -180
   ClientTop       =   330
   ClientWidth     =   10140
   Icon            =   "LocEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5340
      TabIndex        =   67
      Top             =   7260
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Validate"
      Height          =   375
      Left            =   6360
      TabIndex        =   65
      Top             =   7260
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&New"
      Height          =   375
      Left            =   3300
      TabIndex        =   64
      Top             =   7260
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   9060
      TabIndex        =   63
      Top             =   7260
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Search"
      Height          =   315
      Left            =   2460
      TabIndex        =   62
      Top             =   6900
      Width           =   795
   End
   Begin VB.TextBox Text11 
      Height          =   315
      Left            =   60
      TabIndex        =   61
      Top             =   6900
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   57
      Top             =   7260
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   6735
      IntegralHeight  =   0   'False
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   58
      Top             =   120
      Width           =   3195
   End
   Begin VB.Frame Frame1 
      Height          =   7185
      Left            =   3300
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   9
         Left            =   3060
         Picture         =   "LocEdit.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   5895
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   8
         Left            =   3060
         Picture         =   "LocEdit.frx":08EE
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   5580
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   7
         Left            =   3060
         Picture         =   "LocEdit.frx":0C52
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   5265
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   6
         Left            =   3060
         Picture         =   "LocEdit.frx":0FB6
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   4950
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   5
         Left            =   3060
         Picture         =   "LocEdit.frx":131A
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   4635
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   4
         Left            =   3060
         Picture         =   "LocEdit.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   4320
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   3
         Left            =   3060
         Picture         =   "LocEdit.frx":19E2
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   4005
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   2
         Left            =   3060
         Picture         =   "LocEdit.frx":1D46
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3690
         Width           =   285
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   1
         Left            =   3060
         Picture         =   "LocEdit.frx":20AA
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3375
         Width           =   285
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Auto #"
         Height          =   315
         Left            =   5700
         TabIndex        =   66
         Top             =   405
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Height          =   285
         Index           =   0
         Left            =   3060
         Picture         =   "LocEdit.frx":240E
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3060
         Width           =   285
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   9
         Left            =   3555
         TabIndex        =   56
         Text            =   "Text10"
         Top             =   5895
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   8
         Left            =   3555
         TabIndex        =   55
         Text            =   "Text10"
         Top             =   5580
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   7
         Left            =   3555
         TabIndex        =   54
         Text            =   "Text10"
         Top             =   5265
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   6
         Left            =   3555
         TabIndex        =   53
         Text            =   "Text10"
         Top             =   4950
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   5
         Left            =   3555
         TabIndex        =   52
         Text            =   "Text10"
         Top             =   4635
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   4
         Left            =   3555
         TabIndex        =   51
         Text            =   "Text10"
         Top             =   4320
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   3
         Left            =   3555
         TabIndex        =   50
         Text            =   "Text10"
         Top             =   4005
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   2
         Left            =   3555
         TabIndex        =   49
         Text            =   "Text10"
         Top             =   3690
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   1
         Left            =   3555
         TabIndex        =   48
         Text            =   "Text10"
         Top             =   3375
         Width           =   3000
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Index           =   0
         Left            =   3555
         TabIndex        =   47
         Text            =   "Text10"
         Top             =   3060
         Width           =   3000
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   4
         Text            =   "Text9"
         Top             =   405
         Width           =   555
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   3555
         TabIndex        =   41
         Text            =   "Text8"
         Top             =   6570
         Width           =   2985
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   9
         Left            =   2205
         TabIndex        =   38
         Text            =   "Text5"
         Top             =   5895
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   9
         Left            =   945
         TabIndex        =   37
         Text            =   "Text4"
         Top             =   5895
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   9
         Left            =   180
         TabIndex        =   36
         Text            =   "Text3"
         Top             =   5895
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   8
         Left            =   2205
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   5580
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   8
         Left            =   945
         TabIndex        =   34
         Text            =   "Text4"
         Top             =   5580
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   8
         Left            =   180
         TabIndex        =   33
         Text            =   "Text3"
         Top             =   5580
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   7
         Left            =   2205
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   5265
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   7
         Left            =   945
         TabIndex        =   31
         Text            =   "Text4"
         Top             =   5265
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   30
         Text            =   "Text3"
         Top             =   5265
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   5
         Left            =   2205
         TabIndex        =   26
         Text            =   "Text5"
         Top             =   4635
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   5
         Left            =   945
         TabIndex        =   25
         Text            =   "Text4"
         Top             =   4635
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   180
         TabIndex        =   24
         Text            =   "Text3"
         Top             =   4635
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   4
         Left            =   2205
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   4320
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   4
         Left            =   945
         TabIndex        =   22
         Text            =   "Text4"
         Top             =   4320
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Text            =   "Text3"
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   2
         Left            =   2205
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   3690
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   2
         Left            =   945
         TabIndex        =   16
         Text            =   "Text4"
         Top             =   3690
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   3690
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   2205
         TabIndex        =   40
         Text            =   "Text7"
         Top             =   6570
         Width           =   1260
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   180
         TabIndex        =   39
         Text            =   "Text6"
         Top             =   6570
         Width           =   1950
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   6
         Left            =   2205
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   4950
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   6
         Left            =   945
         TabIndex        =   28
         Text            =   "Text4"
         Top             =   4950
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   6
         Left            =   180
         TabIndex        =   27
         Text            =   "Text3"
         Top             =   4950
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   3
         Left            =   2205
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   4005
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   3
         Left            =   945
         TabIndex        =   19
         Text            =   "Text4"
         Top             =   4005
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   4005
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   1
         Left            =   2205
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   3375
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   1
         Left            =   945
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   3375
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   180
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   3375
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Index           =   0
         Left            =   2205
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   3060
         Width           =   795
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   0
         Left            =   945
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   3060
         Width           =   1155
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   3060
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   1575
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "LocEdit.frx":2772
         Top             =   1080
         Width           =   6375
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   405
         Width           =   4815
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3465
         X2              =   3465
         Y1              =   3060
         Y2              =   6165
      End
      Begin VB.Label Label7 
         Caption         =   "No."
         Height          =   195
         Left            =   5040
         TabIndex        =   59
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label11 
         Caption         =   "Flags"
         Height          =   195
         Left            =   3555
         TabIndex        =   45
         Top             =   6345
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Zone"
         Height          =   195
         Left            =   2205
         TabIndex        =   44
         Top             =   6345
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Owner"
         Height          =   195
         Left            =   180
         TabIndex        =   43
         Top             =   6345
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Exit Description (Optional)"
         Height          =   195
         Left            =   3555
         TabIndex        =   42
         Top             =   2835
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Destination"
         Height          =   195
         Left            =   2205
         TabIndex        =   8
         Top             =   2835
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Long Exit"
         Height          =   195
         Left            =   945
         TabIndex        =   7
         Top             =   2835
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Short Exit"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   2835
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "Location Description"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   855
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Location Name"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   2475
      End
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label12"
      Height          =   255
      Left            =   60
      TabIndex        =   46
      Top             =   7380
      Width           =   3195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Changed As Integer
Private Sub Command1_Click()
SaveForm
Changed = 0
End Sub

Private Sub Command2_Click(Index As Integer)
LocLookFor = Trim$(Form1.Text5(Index).Text)
If Val(LocLookFor) < 1 Then Exit Sub
LocLookFor = Right$("00000" + LocLookFor, 5)
For T% = 0 To Form1.List1.ListCount - 1
    LocFound = Left$(Form1.List1.List(T%), 5)
    If LocFound = LocLookFor Then
        Found = 1
        Form1.List1.ListIndex = T%
        T% = Form1.List1.ListCount
        Call List1_Click
        Form1.Text1.SetFocus
    End If
Next T%
If Found = 0 Then
    x = MsgBox("Location Not Found!", vbExclamation, Form1.Caption)
End If
    
End Sub

Private Sub Command3_Click()
Static LastSearch
If LastSearch <> Form1.Text11.Text Then
    Form1.List1.ListIndex = 0
Else
    If Form1.List1.ListIndex < Form1.List1.ListCount - 1 Then
        Form1.List1.ListIndex = Form1.List1.ListIndex + 1
    End If
End If

For T% = Form1.List1.ListIndex To Form1.List1.ListCount - 1
    LocName = UCase$(Mid$(Form1.List1.List(T%), 7))
    If InStr(LocName, UCase$(Form1.Text11.Text)) <> 0 Then
        Found = 1
        Form1.List1.ListIndex = T%
        T% = Form1.List1.ListCount
        Call List1_Click
        LastSearch = Form1.Text11.Text
    End If
Next T%
If Found = 0 Then
    If Form1.List1.ListIndex > 0 Then Form1.List1.ListIndex = Form1.List1.ListIndex - 1
    x = MsgBox("No More Matches", vbInformation, Form1.Caption)
End If


End Sub

Private Sub Command4_Click()
If Changed = 1 Then
    SaveForm
    Changed = 0
End If

End

End Sub

Private Sub Command5_Click()
If Changed = 1 Then
    SaveForm
    Changed = 0
End If

ClearForm

Form1.Text1.SetFocus
End Sub

Private Sub Command6_Click()
Form3.Show
Form3.Text1.Text = ""
End Sub

Private Sub Command7_Click()
Handle = FreeFile
RecChar$ = " "
Open App.Path + "\locations\usedcache.dat" For Binary Access Read As Handle
For T% = 1 To 9999
    Get #Handle, T%, RecChar$
    If RecChar$ = "0" Then FreeLoc = T%: T% = 10000
Next T%
Close #Handle
Form1.Text9.Text = Trim$(Str(FreeLoc))

End Sub

Private Sub Command8_Click()
x = MsgBox("Delete this location?", vbQuestion + vbYesNo, App.Title)
If x = 7 Then Exit Sub

LocName$ = Form1.Text9.Text & ".loc"
Kill App.Path + "\locations\" + LocName$
Handle = FreeFile
Open App.Path + "\locations\usedcache.dat" For Binary Access Write As Handle
Put #Handle, Val(Form1.Text9.Text), "0"
Close #Handle
Form1.List1.RemoveItem Form1.List1.ListIndex
ClearForm
End Sub

Private Sub Form_Load()
Form2.Show
LoadList
Form2.Hide
ClearForm
If Form1.List1.ListCount > 0 Then Form1.List1.ListIndex = 0

End Sub
Sub LoadList()
File$ = Dir(App.Path + "\locations\*.loc")
If File$ <> "" Then
    Do
        FileNo$ = Right$("00000" + Left$(File$, InStr(File$, ".") - 1), 5)
        Open App.Path + "\locations\" + File$ For Input Access Read As #1
        Line Input #1, Ln$
        Ln$ = Mid$(Ln$, 3)
        Form1.List1.AddItem FileNo$ + " " + Ln$
        Form2.Label2.Caption = Ln$
        Form2.Label2.Refresh
        Close #1
        File$ = Dir()
    Loop Until File$ = ""
End If
Form1.Label12.Caption = Trim$(Str$(Form1.List1.ListCount) + " Locations.")

End Sub

Sub ClearForm()
Form1.Text1.Text = ""
Form1.Text2.Text = ""
Form1.Text6.Text = ""
Form1.Text7.Text = ""
Form1.Text8.Text = ""
Form1.Text9.Text = ""

For T% = 0 To 9
Form1.Text3(T%).Text = ""
Form1.Text4(T%).Text = ""
Form1.Text5(T%).Text = ""
Form1.Text10(T%).Text = ""
Next T%

End Sub

Sub LoadLoc(ListID)
Dim ExitInfo() As String
ClearForm
ListEntry$ = Form1.List1.List(ListID)
File$ = Left$(ListEntry$, InStr(ListEntry$, " "))

File$ = Trim$(Str$(Val(File$)))
Form1.Text9.Text = File$
File$ = File$ + ".loc"
ExitOffset = 0

Open App.Path + "\locations\" + File$ For Input Access Read As #1
Do Until EOF(1)
    Line Input #1, Ln$
    LineVal = Val(Left$(Ln$, 1))
    Ln$ = Trim$(Mid$(Ln$, 3))
    Select Case LineVal
    Case 0
        Form1.Text1.Text = Ln$
    Case 1
        Form1.Text2.Text = Form1.Text2.Text + Ln$ + vbCr + vbLf
    Case 2
        If ExitOffset < 10 Then
            Ln$ = Ln$ + ",Null"
            ExitInfo() = Split(Ln$, ",")
            Form1.Text3(ExitOffset).Text = Trim$(ExitInfo(0))
            Form1.Text4(ExitOffset).Text = Trim$(ExitInfo(1))
            Form1.Text5(ExitOffset).Text = Trim$(ExitInfo(2))
            If ExitInfo(3) <> "Null" Then Form1.Text10(ExitOffset).Text = Trim$(ExitInfo(3))
            ExitOffset = ExitOffset + 1
        End If
    Case 3
        Form1.Text6.Text = Ln$
    Case 4
        Form1.Text7.Text = Ln$
    Case 5
        Form1.Text8.Text = Ln$
    Case Else
        MsgBox "Unhandled: " & LineVal & " " & Ln$, vbCritical, Form1.Caption
    End Select
Loop
Close #1

Changed = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Changed = 1 Then SaveForm

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub List1_Click()
If Changed = 1 Then
    SaveForm
    Changed = 0
End If

LoadLoc Form1.List1.ListIndex

End Sub

Private Sub List1_DblClick()
ListEntry$ = Form1.List1.List(Form1.List1.ListIndex)
File$ = Left$(ListEntry$, InStr(ListEntry$, " "))

File$ = Trim$(Str$(Val(File$)))
x = Shell("notepad.exe " + App.Path + "\locations\" + File$ + ".loc", vbNormalFocus)

End Sub

Sub SaveForm()
    Dim Desc() As String
    If Trim$(Form1.Text9.Text) = "" Then
        If Form1.Visible = True Then
            y = MsgBox("Cannot save location, it does not have a location number!", vbExclamation, App.Title)
            Form1.Text9.SetFocus
        End If
        Exit Sub
    End If
    
    File$ = Trim$(Form1.Text9.Text) + ".loc"
    x$ = Dir(App.Path + "\locations\" + File$)
    If x$ <> "" Then
        x = MsgBox(File$ + " already exists, Overwrite?", vbYesNo, Form1.Caption)
        If x = 7 Then Exit Sub
    Else
        ID$ = Right$("00000" + Trim$(Form1.Text9.Text), 5)
        Form1.List1.AddItem ID$ + " " + Trim$(Form1.Text1.Text)
        Form1.List1.Refresh
    End If
    
    Open App.Path + "\locations\" + File$ For Output Access Write As #1
    Print #1, "0 " + Trim$(Form1.Text1.Text)
    x$ = Form1.Text2.Text
    zz$ = ""
    For R% = 1 To Len(x$)
        If Mid$(x$, R%, 1) <> vbLf Then zz$ = zz$ + Mid$(x$, R%, 1)
    Next R%
    Desc() = Split(zz$, vbCr)
    For T% = 0 To UBound(Desc())
        If Trim$(Desc(T%)) <> "" Then
            Print #1, "1 " + Desc(T%)
        End If
    Next T%
    For T% = 0 To 9
        If Trim$(Form1.Text3(T%).Text) <> "" Then
            Ln$ = "2 " + Form1.Text3(T%).Text + "," + Form1.Text4(T%).Text + "," + Form1.Text5(T%).Text
            If Trim$(Form1.Text10(T%).Text) <> "" Then
                Ln$ = Ln$ + "," + Form1.Text10(T%).Text
            End If
            Print #1, Ln$
        End If
    Next T%
    If Trim$(Form1.Text6.Text) <> "" Then Print #1, "3 " + Form1.Text6.Text
    If Trim$(Form1.Text7.Text) <> "" Then Print #1, "4 " + Form1.Text7.Text
    If Trim$(Form1.Text8.Text) <> "" Then Print #1, "5 " + Form1.Text8.Text
    
    Close #1
    
    Handle = FreeFile
    RecChar$ = " "
    
    Open App.Path + "\locations\usedcache.dat" For Binary Access Write As Handle
    Put #Handle, Val(Form1.Text9.Text), "1"
    Close #Handle
    
End Sub

Private Sub Text1_Change()
Changed = 1
End Sub

Private Sub Text10_Change(Index As Integer)
Changed = 1

End Sub

Private Sub Text10_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 1 Then
    Form1.Text10(Index).SelStart = 0
    Form1.Text10(Index).SelLength = Len(Form1.Text10(Index).Text)
    KeyAscii = 0
End If

    
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command3_Click

End Sub

Private Sub Text2_Change()
Changed = 1

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
    Form1.Text2.SelStart = 0
    Form1.Text2.SelLength = Len(Form1.Text2.Text)
    KeyAscii = 0
End If
End Sub

Private Sub Text3_Change(Index As Integer)
Changed = 1

End Sub

Private Sub Text4_Change(Index As Integer)
Changed = 1

End Sub

Private Sub Text5_Change(Index As Integer)
Changed = 1

End Sub

Private Sub Text6_Change()
Changed = 1

End Sub

Private Sub Text7_Change()
Changed = 1

End Sub

Private Sub Text8_Change()
Changed = 1

End Sub
