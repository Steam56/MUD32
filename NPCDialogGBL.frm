VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "NPCDialogGBL.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "NPCDialogGBL.frx":058A
      Top             =   0
      Width           =   3675
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form4.Visible = False
Cancel = 1

End Sub

Private Sub Form_Resize()
Form4.Text1.Width = Form4.ScaleWidth
Form4.Text1.Height = Form4.ScaleHeight

End Sub

