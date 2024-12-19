VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ChangePlayerPass Form1.Text3.Text, Form1.Text1.Text, Form1.Text2.Text, 0
End Sub

Function Crypt(Pass$)
For T% = 1 To Len(Pass$)
    x% = Asc(Mid$(Pass$, T%, 1))
    y% = x% Xor (27 + T%)
    NewPass$ = NewPass$ + Chr$(y%)
Next T%
Crypt = NewPass$

End Function

Sub ChangePlayerPass(PlrName, uOldPass$, uNewPass$, PreCheck)
Dim FileBuff(50) As String
OldPass = Crypt(uOldPass$)
NewPass = Crypt(uNewPass$)

PlayerFile = App.Path + "\players\" + PlrName + ".plr"
x = Dir(PlayerFile)
If x = "" Then
    Debug.Print "Player does not exist"
    Exit Sub
End If

Handle = FreeFile
Counter = 0
Open PlayerFile For Input Access Read As #Handle
Do Until EOF(Handle)
    Line Input #Handle, Ln$
    FileBuff(Counter) = Ln$
    Counter = Counter + 1
Loop
Close #Handle
If FileBuff(1) <> OldPass And PreCheck = 1 Then
    Debug.Print "Old Password wrong"
    Exit Sub
End If
FileBuff(1) = NewPass
Handle = FreeFile
Counter = 0
Open PlayerFile For Output Access Write As #Handle
Do Until FileBuff(Counter) = ""
    Print #Handle, FileBuff(Counter)
    Counter = Counter + 1
Loop
Close #Handle
Debug.Print "Password changed."

End Sub
