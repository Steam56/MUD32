VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MUD32 AutoMapper"
   ClientHeight    =   7470
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   9270
   Icon            =   "AutoMap.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7005
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1185
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Loc IDs"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   45
         Top             =   6180
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         TabIndex        =   43
         Text            =   "1"
         Top             =   6480
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Single Zone"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   42
         Top             =   5880
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ê"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   40
         Top             =   5400
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   38
         Top             =   5400
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Map!"
         Height          =   285
         Left            =   45
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   90
         TabIndex        =   21
         Text            =   "1"
         Top             =   45
         Width           =   1050
      End
      Begin VB.Label Label7 
         Caption         =   "Scale"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   6540
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   240
         Left            =   420
         TabIndex        =   41
         Top             =   5460
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   60
         TabIndex        =   39
         Top             =   5160
         Width           =   1050
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   15
         Left            =   45
         TabIndex        =   37
         Top             =   4725
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   14
         Left            =   45
         TabIndex        =   36
         Top             =   4455
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   13
         Left            =   45
         TabIndex        =   35
         Top             =   4185
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   12
         Left            =   45
         TabIndex        =   34
         Top             =   3915
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   11
         Left            =   45
         TabIndex        =   33
         Top             =   3645
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   10
         Left            =   45
         TabIndex        =   32
         Top             =   3375
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   9
         Left            =   45
         TabIndex        =   31
         Top             =   3105
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   8
         Left            =   45
         TabIndex        =   30
         Top             =   2835
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   7
         Left            =   45
         TabIndex        =   29
         Top             =   2565
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   6
         Left            =   45
         TabIndex        =   28
         Top             =   2295
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   5
         Left            =   45
         TabIndex        =   27
         Top             =   2025
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   4
         Left            =   45
         TabIndex        =   26
         Top             =   1755
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   3
         Left            =   45
         TabIndex        =   25
         Top             =   1485
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   2
         Left            =   45
         TabIndex        =   24
         Top             =   1215
         Width           =   195
      End
      Begin VB.Label Label4 
         Height          =   240
         Index           =   1
         Left            =   45
         TabIndex        =   23
         Top             =   945
         Width           =   195
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000002&
         Caption         =   "Legend"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   45
         TabIndex        =   20
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   15
         Left            =   315
         TabIndex        =   19
         Top             =   4725
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   14
         Left            =   315
         TabIndex        =   18
         Top             =   4455
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   13
         Left            =   315
         TabIndex        =   17
         Top             =   4185
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   12
         Left            =   315
         TabIndex        =   16
         Top             =   3915
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   11
         Left            =   315
         TabIndex        =   15
         Top             =   3645
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   10
         Left            =   315
         TabIndex        =   14
         Top             =   3375
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   9
         Left            =   315
         TabIndex        =   13
         Top             =   3105
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   8
         Left            =   315
         TabIndex        =   12
         Top             =   2835
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   7
         Left            =   315
         TabIndex        =   11
         Top             =   2565
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   6
         Left            =   315
         TabIndex        =   10
         Top             =   2295
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   5
         Left            =   315
         TabIndex        =   9
         Top             =   2025
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   4
         Left            =   315
         TabIndex        =   8
         Top             =   1755
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   3
         Left            =   315
         TabIndex        =   7
         Top             =   1485
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   2
         Left            =   315
         TabIndex        =   6
         Top             =   1215
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   5
         Top             =   945
         Width           =   825
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      LargeChange     =   10
      Left            =   3375
      Max             =   500
      TabIndex        =   2
      Top             =   6300
      Width           =   3210
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2040
      LargeChange     =   10
      Left            =   8235
      Max             =   500
      TabIndex        =   1
      Top             =   2160
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5235
      Left            =   1215
      ScaleHeight     =   5175
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "AutoMapper"
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   7200
      Width           =   1230
   End
   Begin VB.Menu Mnu_File 
      Caption         =   "&File"
      Begin VB.Menu Mnu_BMPSave 
         Caption         =   "&Save BMP"
      End
      Begin VB.Menu Mnu_CLose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RoomArray(10, 500, 500, 11) As Integer
Dim ZoneList(100) As String
Dim LocList() As Integer
Dim CurrentLocX As Integer
Dim CurrentLocY As Integer
Dim LoadRoom As Integer
Dim ZAxis As Integer
Dim LocPath As String


Private Sub Command1_Click()
Form1.Command1.Enabled = False

'Setup map area
CurrentLocX = 250
CurrentLocY = 250
Form1.HScroll1.Value = CurrentLocX - 8
Form1.VScroll1.Value = CurrentLocY - 8
ZAxis = 1
Form1.Label6.Caption = ZAxis
'Form1.Picture1.Picture = LoadPicture("c:\windows\zapotec.bmp")

'clear zonelist
For T% = 0 To 100: ZoneList(T%) = "": Next T%
For T% = 1 To 15
    Form1.Label2(T%).Caption = ""
Next T%


'Zero Room Array
For H% = 1 To 10
    For X% = 0 To 500
        For Y% = 0 To 500
            For Z% = 0 To 11
                RoomArray(H%, X%, Y%, Z%) = 0
            Next Z%
        Next Y%
    Next X%
Next H%

ReDim LocList(10000)

LoadExits Val(Form1.Text1.Text)

WalkMap

Form1.Command1.Enabled = True

End Sub

Private Sub Command2_Click(Index As Integer)
If Index = 0 Then
    Form1.Label6.Caption = Form1.Label6.Caption + 1
    If Form1.Label6.Caption > 10 Then Form1.Label6.Caption = 10
Else
    Form1.Label6.Caption = Form1.Label6.Caption - 1
    If Form1.Label6.Caption < 1 Then Form1.Label6.Caption = 1
End If
ZAxis = Form1.Label6.Caption
RedrawMap
End Sub

Private Sub Form_Load()

LocPath = Left(App.Path, InStrRev(App.Path, "\")) + "locations\"

For T% = 1 To 15
    Form1.Label2(T%).Caption = ""
Next T%
Form1.Show
'Form1.WindowState = 2







End Sub
Sub RedrawMap()
Form1.Picture1.Line (0, 0)-(Form1.Picture1.Width, Form1.Picture1.Height), Form1.Picture1.BackColor, BF

For Y% = 1 To 500
    For X% = 1 To 500
        If RoomArray(ZAxis, X%, Y%, 0) <> 0 Then
            If Form1.Check1.Value = 1 Then
                If RoomArray(ZAxis, X%, Y%, 11) = 1 Then
                    DrawLoc X%, Y%, RoomArray(ZAxis, X%, Y%, 1), RoomArray(ZAxis, X%, Y%, 2), RoomArray(ZAxis, X%, Y%, 3), RoomArray(ZAxis, X%, Y%, 4), RoomArray(ZAxis, X%, Y%, 5), RoomArray(ZAxis, X%, Y%, 6), RoomArray(ZAxis, X%, Y%, 7), RoomArray(ZAxis, X%, Y%, 8), RoomArray(ZAxis, X%, Y%, 9), RoomArray(ZAxis, X%, Y%, 10), RoomArray(ZAxis, X%, Y%, 0)
                End If
                
            Else
                DrawLoc X%, Y%, RoomArray(ZAxis, X%, Y%, 1), RoomArray(ZAxis, X%, Y%, 2), RoomArray(ZAxis, X%, Y%, 3), RoomArray(ZAxis, X%, Y%, 4), RoomArray(ZAxis, X%, Y%, 5), RoomArray(ZAxis, X%, Y%, 6), RoomArray(ZAxis, X%, Y%, 7), RoomArray(ZAxis, X%, Y%, 8), RoomArray(ZAxis, X%, Y%, 9), RoomArray(ZAxis, X%, Y%, 10), RoomArray(ZAxis, X%, Y%, 0)
            End If
        End If
    Next X%
Next Y%

Form1.Label1.Caption = "x" + Trim$(Str$(Form1.HScroll1.Value)) + " y" + Trim$(Str$(Form1.VScroll1.Value)) + ""

End Sub
Sub DrawLoc(X, Y, N, S, E, W, NE, SE, NW, SW, U, D, ID)
uScale = Abs(300 * Form1.Text2.Text)
uScaleHalf = (uScale / 2)
uPoint = uScale / 10

RoomFill = QBColor(RoomArray(ZAxis, X, Y, 11))
RoomBorder = QBColor(0)
LinkCol = QBColor(0)

uX = ((X - Form1.HScroll1.Value) * 2) * uScale
uY = ((Y - Form1.VScroll1.Value) * 2) * uScale

'object.Line [Step] (x1, y1) [Step] - (x2, y2), [color], [B][F]

Form1.Picture1.Line (uX, uY)-(uX + uScale, uY + uScale), RoomFill, BF
Form1.Picture1.Line (uX, uY)-(uX + uScale, uY + uScale), RoomBorder, B

If N <> 0 Then Form1.Picture1.Line (uX + uScaleHalf, uY)-(uX + uScaleHalf, uY - (uScaleHalf + 10)), LinkCol
If S <> 0 Then Form1.Picture1.Line (uX + uScaleHalf, uY + uScale)-(uX + uScaleHalf, uY + uScale + uScaleHalf + 10), LinkCol
If E <> 0 Then Form1.Picture1.Line (uX + uScale, uY + uScaleHalf)-(uX + uScale + uScaleHalf + 10, uY + uScaleHalf), LinkCol
If W <> 0 Then Form1.Picture1.Line (uX, uY + uScaleHalf)-(uX - (uScaleHalf + 10), uY + uScaleHalf), LinkCol

If NE <> 0 Then Form1.Picture1.Line (uX + uScale, uY)-(uX + uScale + uScaleHalf + 10, uY - (uScaleHalf + 20)), LinkCol
If SE <> 0 Then Form1.Picture1.Line (uX + uScale, uY + uScale)-(uX + uScale + uScaleHalf + 10, uY + uScale + uScaleHalf + 10), LinkCol
If SW <> 0 Then Form1.Picture1.Line (uX, uY + uScale)-(uX - (uScaleHalf + 10), uY + uScale + uScaleHalf + 10), LinkCol
If NW <> 0 Then Form1.Picture1.Line (uX, uY)-(uX - (uScaleHalf + 10), uY - (uScaleHalf + 10)), LinkCol
If U <> 0 Then
    
    Form1.Picture1.Line (uX + (uPoint * 2.5), uY + (uPoint * 9))-(uX + (uPoint * 2.5), uY + (uPoint * 5)), RoomBorder
    Form1.Picture1.Line (uX + (uPoint * 1.5), uY + (uPoint * 7))-(uX + (uPoint * 2.5), uY + (uPoint * 5)), RoomBorder
    Form1.Picture1.Line (uX + (uPoint * 3.5), uY + (uPoint * 7))-(uX + (uPoint * 2.5), uY + (uPoint * 5)), RoomBorder
    
End If
If D <> 0 Then
   
    Form1.Picture1.Line (uX + (uPoint * 7.5), uY + (uPoint * 9))-(uX + (uPoint * 7.5), uY + (uPoint * 5)), RoomBorder
    Form1.Picture1.Line (uX + (uPoint * 6.5), uY + (uPoint * 7))-(uX + (uPoint * 7.5), uY + (uPoint * 9)), RoomBorder
    Form1.Picture1.Line (uX + (uPoint * 8.5), uY + (uPoint * 7))-(uX + (uPoint * 7.5), uY + (uPoint * 9)), RoomBorder
    
End If
If uScale > 200 And Form1.Check2.Value = 1 Then
    Form1.Picture1.CurrentX = uX
    Form1.Picture1.CurrentY = uY
    Form1.Picture1.Print ID
End If



End Sub

Private Sub Form_Resize()
Form1.Frame1.Top = 0
Form1.Frame1.Left = 0
Form1.VScroll1.Left = Form1.ScaleWidth - Form1.VScroll1.Width
Form1.VScroll1.Top = 0
Form1.VScroll1.Height = Abs(Form1.ScaleHeight - Form1.HScroll1.Height)
Form1.Picture1.Left = Form1.Frame1.Width
Form1.Picture1.Width = Abs(Form1.ScaleWidth - Form1.VScroll1.Width - Form1.Frame1.Width)
Form1.Picture1.Height = Abs(Form1.ScaleHeight - Form1.HScroll1.Height)
Form1.HScroll1.Left = Form1.Picture1.Left
Form1.HScroll1.Top = Abs(Form1.ScaleHeight - Form1.HScroll1.Height)
Form1.HScroll1.Width = Form1.Picture1.Width
Form1.Label1.Left = 0
Form1.Label1.Top = Form1.HScroll1.Top + 20

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll1_Change()
RedrawMap
End Sub

Private Sub Mnu_BMPSave_Click()
SavePicture Form1.Picture1.Picture, App.Path + "\untitled.bmp"

End Sub

Private Sub Mnu_CLose_Click()
End
End Sub

Private Sub VScroll1_Change()
RedrawMap
End Sub

Sub LoadExits(LocNo)
If LocList(Val(LocNo)) = 1 Then Exit Sub 'if we've done this loc already then exit
X = Dir(LocPath + Trim$(Str$(LocNo)) + ".loc")
Debug.Print ".." + LocPath + Trim$(Str$(LocNo)) + ".loc"
If X = "" Then Exit Sub
Debug.Print "Opening: " + Trim$(Str$(LocNo)) + ".loc"
Open LocPath + Trim$(Str$(LocNo)) + ".loc" For Input Access Read As #1
Do Until EOF(1)
Line Input #1, Ln$
If LocNo = 7 Then Debug.Print ">> " + Ln$

If Left$(Ln$, 1) = "2" Then
    Debug.Print Ln$
    Ln$ = Ln$ + " , "
    Ln$ = Mid$(Ln$, 3)
    I% = InStr(Ln$, ",")
    SExit = UCase$(Left$(Ln$, I% - 1))
    Ln$ = Mid$(Ln$, I% + 1)
    I% = InStr(Ln$, ",")
    Ln$ = Mid$(Ln$, I% + 1)
    I% = InStr(Ln$, ",")
    ExitLoc = Left$(Ln$, I% - 1)
    Debug.Print SExit, ExitLoc
    If SExit = "N" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 1) = Val(ExitLoc)
    If SExit = "S" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 2) = Val(ExitLoc)
    If SExit = "E" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 3) = Val(ExitLoc)
    If SExit = "W" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 4) = Val(ExitLoc)
    If SExit = "NE" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 5) = Val(ExitLoc)
    If SExit = "SE" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 6) = Val(ExitLoc)
    If SExit = "NW" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 7) = Val(ExitLoc)
    If SExit = "SW" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 8) = Val(ExitLoc)
    If SExit = "U" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 9) = Val(ExitLoc)
    If SExit = "D" Then RoomArray(ZAxis, CurrentLocX, CurrentLocY, 10) = Val(ExitLoc)
    RoomArray(ZAxis, CurrentLocX, CurrentLocY, 0) = Val(LocNo)
    LocList(Val(LocNo)) = 1
    
End If

If Left$(Ln$, 2) = "4 " Then
    ZName$ = Trim(Mid$(Ln$ + Space$(20), 3))
    Zindex = 1
    For T% = 1 To 100
        If ZoneList(T%) = ZName$ Then
            Zindex = T%
            T% = 101
        ElseIf ZoneList(T%) = "" Then
            Zindex = T%
            Debug.Print LocNo, ZName$
            T% = 101
        End If
    Next T%
    ZoneList(Zindex) = ZName$
    Form1.Label2(Zindex).Caption = ZName$
    Form1.Label4(Zindex).BackColor = QBColor(Zindex)
    
    RoomArray(ZAxis, CurrentLocX, CurrentLocY, 11) = Zindex
End If

Loop
Close #1



LoadRoom = 1


End Sub


Sub WalkMap()
Form2.Show
SHW = Form2.Shape1.Width
Form2.Shape1.Width = 0

Do
LoadRoom = 0
For Z% = 1 To 10

    For Y% = 2 To 499
        For X% = 2 To 499
            CurrentLocX = X%
            CurrentLocY = Y%
            ZAxis = Z%
            I% = RoomArray(Z%, X%, Y%, 0)
            If I% <> 0 Then
                I% = RoomArray(Z%, X%, Y%, 1)
                If I% <> 0 And RoomArray(Z%, X%, Y% - 1, 0) = 0 Then
                    CurrentLocX = X%
                    CurrentLocY = Y% - 1
                    LoadExits I%
                End If
                I% = RoomArray(Z%, X%, Y%, 2)
                If I% <> 0 And RoomArray(Z%, X%, Y% + 1, 0) = 0 Then
                    CurrentLocX = X%
                    CurrentLocY = Y% + 1
                    LoadExits I%
                End If
                I% = RoomArray(Z%, X%, Y%, 3)
                If I% <> 0 And RoomArray(Z%, X% + 1, Y%, 0) = 0 Then
                    CurrentLocX = X% + 1
                    CurrentLocY = Y%
                    LoadExits I%
                End If
                
                I% = RoomArray(Z%, X%, Y%, 4)
                If I% <> 0 And RoomArray(Z%, X% - 1, Y%, 0) = 0 Then
                    CurrentLocX = X% - 1
                    CurrentLocY = Y%
                    LoadExits I%
                End If
                
                I% = RoomArray(Z%, X%, Y%, 5)
                If I% <> 0 And RoomArray(Z%, X% + 1, Y% - 1, 0) = 0 Then
                    CurrentLocX = X% + 1
                    CurrentLocY = Y% - 1
                    LoadExits I%
                End If
                I% = RoomArray(Z%, X%, Y%, 6)
                If I% <> 0 And RoomArray(Z%, X% + 1, Y% + 1, 0) = 0 Then
                    CurrentLocX = X% + 1
                    CurrentLocY = Y% + 1
                    LoadExits I%
                End If
                I% = RoomArray(Z%, X%, Y%, 7)
                If I% <> 0 And RoomArray(Z%, X% - 1, Y% - 1, 0) = 0 Then
                    CurrentLocX = X% - 1
                    CurrentLocY = Y% - 1
                    LoadExits I%
                End If
                I% = RoomArray(Z%, X%, Y%, 9)
                If I% <> 0 And RoomArray(Z% + 1, X%, Y%, 0) = 0 Then
                    CurrentLocX = X%
                    CurrentLocY = Y%
                    ZAxis = Z% + 1
                    LoadExits I%
                End If
                'I% = RoomArray(Z%, X%, Y%, 10)
                'If I% <> 0 And RoomArray(Z% - 1, X%, Y%, 0) = 0 Then
                '    CurrentLocX = X%
                '    CurrentLocY = Y%
                '    ZAxis = Z% - 1
                '    LoadExits I%
                'End If
                
            End If
            
        Next X%
    Next Y%
    Form2.Shape1.Width = (SHW / 10) * Z%
Next Z%
ZAxis = 1
RedrawMap
DoEvents

Loop Until LoadRoom = 0
Form2.Hide

End Sub
