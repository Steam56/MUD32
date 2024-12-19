VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3555
   ControlBox      =   0   'False
   Icon            =   "LocEditSplash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Locations..."
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   780
      Width           =   2475
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Editor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   3315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MUD32"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "LocEditSplash.frx":0882
      Stretch         =   -1  'True
      Top             =   180
      Width           =   960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
