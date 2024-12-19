VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUD32 Notify"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4860
   Icon            =   "MUD2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   660
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2580
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2760
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   120
      Top             =   3120
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1515
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2160
      Picture         =   "MUD2.frx":014A
      Top             =   2580
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1740
      Picture         =   "MUD2.frx":0294
      Top             =   2580
      Width           =   240
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
RespawnItems
NPC_Respawn
BellTolls

End Sub

Private Sub Timer2_Timer()
NPC_Action

If Form1.Visible = True Then
    Form1.Label6.Caption = GetUptimeConsole
End If

End Sub
