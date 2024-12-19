VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Container for SystemTray Icon and Menu"
   ClientHeight    =   3210
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6270
   Icon            =   "MUD_SysTray.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3210
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim Result As Long
  Dim lMessage As Long

  ' the value of X will vary depending upon the
  ' scalemode setting
  If Me.ScaleMode = vbPixels Then
    lMessage = x
  Else
    lMessage = x / Screen.TwipsPerPixelX
  End If

  If WM_RBUTTONUP = lMessage Then
    Me.PopupMenu Form1.Mnu_MUD
  End If
  If WM_LBUTTONDBLCLK = lMessage Then
    Call Form1.Mnu_HideConsole_Click
    
  End If
  

End Sub
