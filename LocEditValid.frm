VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Location Validator"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "LocEditValid.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1140
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   315
      Left            =   5880
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Validate"
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Text1.Text = ""
Form3.File1.Refresh
Form3.List1.Clear
IsErrors = 0
vPrint "Validating..." + vbCrLf
Form3.File1.Path = App.Path + "\locations"
Form3.File1.Pattern = "*.loc"
For T% = 0 To Form3.File1.ListCount - 1
    Form3.List1.AddItem Form3.File1.List(T%)
Next T%


For T% = 0 To Form3.File1.ListCount - 1
    CurrentLoc = Form3.File1.List(T%)
    'vPrint "Checking " + CurrentLoc
    Validate CurrentLoc
Next T%

For S% = 0 To Form3.List1.ListCount - 1
    vPrint "WARNING: Location " + Left(Form3.List1.List(S%), InStr(Form3.List1.List(S%), ".") - 1) + " is not connected to any other location."
    IsErrors = 1
Next S%
If IsErrors = 0 Then vPrint "No Errors Found."
vPrint vbCrLf + "Validation Complete." + Str(Form3.File1.ListCount) + " locations."

  
    

End Sub

Sub vPrint(uText)
Form3.Text1.Text = Form3.Text1.Text + uText + vbCrLf
Form3.Text1.SelStart = Len(Form3.Text1.Text)
End Sub

Sub Validate(CurLoc)
Open App.Path + "\locations\" + CurLoc For Input Access Read As #7
Do Until EOF(7)
    Line Input #7, Ln$
    uLine = Trim(Ln$)
    If Left(uLine, 1) = "0" Then uName = Mid(uLine, 3)
    If Left(uLine, 1) = "2" Then 'It's an exit line
        uLine = Trim(Mid(uLine, 2))
        result = CheckExit(uLine)
        If result = 1 Then uError = "has bad exit. [" + uLine + "]"
        If result = 2 Then uError = "exit line is invalid [" + uLine + "]"
        If result > 0 Then
            vPrint "ERROR: Location " + Left(CurLoc, InStr(CurLoc, ".") - 1) + ", '" + uName + "' " + uError
            IsErrors = 1
        End If
    End If
Loop
Close #7


End Sub
Function CheckExit(uLine)
Dim Bits() As String
Found = 0
CheckExit = 0
Bits() = Split(uLine, ",")
Chkloc = Trim(Bits(2))
'vPrint Chkloc + " " + uLine
If Chkloc <> "" Then
    For R% = 0 To Form3.File1.ListCount
        If UCase(Form3.File1.List(R%)) = Chkloc + ".LOC" Then
            Found = 1
            R% = Form3.File1.ListCount + 1
            FndIndex = -1
            For S% = 0 To Form3.List1.ListCount - 1
                If UCase(Form3.List1.List(S%)) = Chkloc + ".LOC" Then FndIndex = S%: S% = Form1.List1.ListCount + 1
            Next S%
            If FndIndex <> -1 Then Form3.List1.RemoveItem FndIndex
        End If
    Next R%
    If Found = 0 Then CheckExit = 1
Else
    CheckExit = 2
End If
        
End Function

Private Sub Command2_Click()
Form3.Hide

End Sub

