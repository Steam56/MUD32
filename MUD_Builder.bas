Attribute VB_Name = "Module3"
Sub Builder(ID, Param$)
OldParam$ = Param$
Dim PArray(10) As String
HelpMsg$ = "For help on the building command type HELP BUILDING."

If Val(UserData(ID, 9)) = 0 Then
    Exit Sub
End If

If Param$ = "" Or InStr(Param$, " ") = 0 Then
    SendData ID, HelpMsg$
    Exit Sub
End If
Element = 0
Do
    I% = InStr(Param$, " ")
    If I% <> 0 Then
        PArray(Element) = Trim$(Left$(Param$, InStr(Param$, " ")))
        Param$ = Trim$(Mid$(Param$, InStr(Param$, " ")))
        Element = Element + 1
    Else
        If Param$ <> "" Then PArray(Element) = Param$
    End If
Loop Until I% = 0 Or Element = 11
       
'For T% = 0 To 10: Debug.Print PArray(T%): Next T%

Keyword$ = UCase$(PArray(0))

If Keyword$ = "NEWLOC" Then
    LocName$ = ""
    For T% = 1 To 10
        LocName$ = LocName$ + PArray(T%) + " "
    Next T%
    LocName$ = Trim$(LocName$)
    Form2.File1.Path = App.Path + "\locations"
    Form2.File1.Pattern = "*.loc"
    Handle = FreeFile
    RecChar$ = " "
    Open App.Path + "\locations\usedcache.dat" For Binary Access Read As Handle
    For T% = 1 To 9999
        Get #Handle, T%, RecChar$
        If RecChar$ = "0" Then FreeLoc = T%: T% = 10000
    Next T%
    Close #Handle
    'LocFname$ = Trim$(Str$(FreeLoc)) + ".loc"
    'Handle = FreeFile
    'Open App.Path + "\locations\" + LocFname$ For Output Access Write As Handle
    'Print #Handle, "0 " + LocName
   ' Print #Handle, "3 " + UserData(ID, 0)
   ' Close #Handle
    Call CreateLoc(FreeLoc, LocName$, UserData(ID, 0))
    Call ToggleLocCache(FreeLoc, "1")
    SendData ID, "Location" + Str$(FreeLoc) + " created, with name '" + LocName + "'."
    Exit Sub
End If

If Keyword$ = "REMLOC" Then
    FullLocName$ = App.Path + "\locations\" + Trim$(PArray(1)) + ".loc"
    LocNo = Val(PArray(1))
    If LocNo = Val(UserData(ID, 10)) Then
        SendData ID, "You cannot remove a location while you are standing in it."
        Exit Sub
    End If
    If LocNo < 1 Or LocNo > 99999 Then
        SendData ID, "Invalid location number."
        Exit Sub
    End If
    If IfExist(FullLocName$) Then
        Owner$ = GetLocEntry(LocNo, "3")
        If Owner$ <> UserData(ID, 0) And Val(UserData(ID, 9)) < 2 Then
            SendData ID, "Location" + Str$(LocNo) + " does not belong to you, therefore you cannot delete it."
            Exit Sub
        End If
        Kill FullLocName$
        Call ToggleLocCache(LocNo, "0")
        For T% = 0 To 64
            X% = Val(UserData(T%, 10))
            If X% = LocNo Then
                SendData T%, "Suddenly the world about you shimmers and fades out and you find yourself someplace else."
                SayPrompt ID, ""
                UserData(T%, 10) = "1"
            End If
        Next T%
        SendData ID, "Location" + Str$(LocNo) + " deleted."
    Else
        SendData ID, "Cannot delete location" + Str$(LocNo) + ", it doesn't exist."
    End If
    Exit Sub
End If

If Keyword$ = "LOCINFO" Then
    FullLocName$ = App.Path + "\locations\" + Trim$(PArray(1)) + ".loc"
    LocNo = Val(PArray(1))
    If LocNo < 1 Or LocNo > 99999 Then
        SendData ID, "Invalid location number."
        Exit Sub
    End If
    If IfExist(FullLocName$) Then
        LocName$ = GetLocEntry(LocNo, "0")
        LocDesc$ = GetLocEntry(LocNo, "1")
        Owner$ = GetLocEntry(LocNo, "3")
        LocZone$ = GetLocEntry(LocNo, "4")
        LocFl$ = GetLocEntry(LocNo, "5")
        SendData ID, "Location" + Str$(LocNo)
        SendData ID, "Name: " + LocName$
        SendData ID, "Description: " + LocDesc$
        SendData ID, "Owner: " + Owner$
        SendData ID, "Zone: " + LocZone$
        SendData ID, "Flags: " + LocFl$
    Else
        SendData ID, "Location" + Str$(LocNo) + " doesn't exist."
    End If
    Exit Sub
End If

If Keyword$ = "LOCEXIT" Then
    FullLocName$ = App.Path + "\locations\" + Trim$(PArray(1)) + ".loc"
    LocNo = Val(PArray(1))
    If LocNo < 1 Or LocNo > 99999 Then
        SendData ID, "Invalid location number."
        Exit Sub
    End If
    If IfExist(FullLocName$) Then
        Handle = FreeFile
        Open FullLocName$ For Input Access Read As #Handle
        Do Until EOF(Handle)
            Line Input #Handle, Ln$
            If Left$(Ln$, 1) = "2" Then
                SendData ID, Mid$(Ln$, 3)
            End If
        Loop
        Close #Handle
    Else
        SendData ID, "Location" + Str$(LocNo) + " doesn't exist."
    End If
    Exit Sub
End If

If Left$(Keyword$, 3) = "LOC" Then
If Keyword$ = "LOCNAME" Then Entry$ = "0": EntryMsg$ = "Name"
If Keyword$ = "LOCDESC" Then Entry$ = "1": EntryMsg$ = "Description"
If Keyword$ = "LOCOWN" And Val(UserData(ID, 9)) > 1 Then Entry$ = "3": EntryMsg$ = "Owner"
If Keyword$ = "LOCZONE" Then Entry$ = "4": EntryMsg$ = "Zone"
If Keyword$ = "LOCFLAG" Then Entry$ = "5": EntryMsg$ = "Flags"
If Entry$ = "" Then
    SendData ID, HelpMsg$
    Exit Sub
End If
I% = InStr(OldParam$, PArray(2))
If PArray(2) <> "" Then
    NewParam$ = Mid$(OldParam$, I%)
Else
    NewParam$ = ""
End If

FullLocName$ = App.Path + "\locations\" + Trim$(PArray(1)) + ".loc"
    LocNo = Val(PArray(1))
    If LocNo < 1 Or LocNo > 99999 Then
        SendData ID, "Invalid location number."
        Exit Sub
    End If
    If IfExist(FullLocName$) Then
        Owner$ = GetLocEntry(LocNo, "3")
        If Owner$ <> UserData(ID, 0) And Val(UserData(ID, 9)) < 2 Then
            SendData ID, "Location" + Str$(LocNo) + " does not belong to you, therefore you cannot change it."
            Exit Sub
        End If
        Call WriteLoc(LocNo, Entry$, NewParam$)
        SendData ID, EntryMsg$ + " for location" + Str$(LocNo) + " changed to '" + NewParam$ + "'."
    Else
        SendData ID, "Location" + Str$(LocNo) + " doesn't exist."
    End If
    Exit Sub
End If
    
If Keyword$ = "DIG" Then
    DigDir$ = UCase$(PArray(1))
    If DigDir$ = "N" Then DigDirLong$ = "North": DigBackDir$ = "S": DigBackDirLong$ = "South"
    If DigDir$ = "S" Then DigDirLong$ = "South": DigBackDir$ = "N": DigBackDirLong$ = "North"
    If DigDir$ = "E" Then DigDirLong$ = "East": DigBackDir$ = "W": DigBackDirLong$ = "West"
    If DigDir$ = "W" Then DigDirLong$ = "West": DigBackDir$ = "E": DigBackDirLong$ = "East"
    If DigDir$ = "NE" Then DigDirLong$ = "NorthEast": DigBackDir$ = "SW": DigBackDirLong$ = "SouthWest"
    If DigDir$ = "NW" Then DigDirLong$ = "NorthWest": DigBackDir$ = "SE": DigBackDirLong$ = "SouthEast"
    If DigDir$ = "SE" Then DigDirLong$ = "SouthEast": DigBackDir$ = "NW": DigBackDirLong$ = "NorthWest"
    If DigDir$ = "SW" Then DigDirLong$ = "SouthWest": DigBackDir$ = "NE": DigBackDirLong$ = "NorthEast"
    If DigDir$ = "U" Then DigDirLong$ = "Up": DigBackDir$ = "D": DigBackDirLong$ = "Down"
    If DigDir$ = "D" Then DigDirLong$ = "Down": DigBackDir$ = "U": DigBackDirLong$ = "Up"
    If DigDir$ = "O" Then DigDirLong$ = "Out": DigBackDir$ = "": DigBackDirLong$ = ""
    
    FullFromLocName$ = App.Path + "\locations\" + Trim$(PArray(2)) + ".loc"
    FullToLocName$ = App.Path + "\locations\" + Trim$(PArray(3)) + ".loc"
    FromLocNo = Val(PArray(2))
    ToLocNo = Val(PArray(3))
    
    If FromLocNo < 1 Or FromLocNo > 99999 Then
        SendData ID, "Invalid FROM location number."
        Exit Sub
    End If
    If ToLocNo < 1 Or ToLocNo > 99999 Then
        SendData ID, "Invalid TO location number."
        Exit Sub
    End If
    If IfExist(FullFromLocName$) Then
        FirstExit$ = DigDir$ + ", " + DigDirLong$ + ", " + PArray(3)
        SecondExit$ = DigBackDir$ + ", " + DigBackDirLong$ + ", " + PArray(2)
        
        Owner$ = GetLocEntry(FromLocNo, "3")
        If Owner$ <> UserData(ID, 0) And Val(UserData(ID, 9)) < 2 Then
            SendData ID, "Location" + Str$(FromLocNo) + " does not belong to you, therefore you cannot change it."
            Exit Sub
        Else
            Call WriteLocExit(FromLocNo, FirstExit$)
            SendData ID, "Exit " + DigDirLong$ + " from location" + Str$(FromLocNo) + " to location" + Str$(ToLocNo) + " created."
            PlayerLoc = Val(UserData(ID, 10))
            If FromLocNo = PlayerLoc Then GetExits ID, PlayerLoc
                    
        End If
        If IfExist(FullToLocName$) Then
    
            If DigBackDir$ <> "" Then
                Owner$ = GetLocEntry(ToLocNo, "3")
                If Owner$ <> UserData(ID, 0) And Val(UserData(ID, 9)) < 2 Then
                    SendData ID, "Location" + Str$(ToLocNo) + " does not belong to you, therefore you cannot change it. Ask the owner of location" + Str$(ToLocNo) + " (" + Owner$ + ") to DIG into location" + Str$(FromLocNo) + "."
                    Exit Sub
                Else
                    Call WriteLocExit(ToLocNo, SecondExit$)
                    SendData ID, "Exit " + DigBackDirLong$ + " from location" + Str$(ToLocNo) + " to location" + Str$(FromLocNo) + " created."
                    PlayerLoc = Val(UserData(ID, 10))
                    If ToLocNo = PlayerLoc Then GetExits ID, PlayerLoc
                End If
            End If
        Else
            SendData ID, "Location" + Str$(ToLocNo) + " doesn't exist."
        End If
    Else
        SendData ID, "Location" + Str$(FromLocNo) + " doesn't exist."
    End If
    Exit Sub
End If

If Keyword$ = "UNDIG" Then
    DigDir$ = UCase$(PArray(1))
    LocNo = Val(PArray(2))
    FullLocName$ = App.Path + "\locations\" + Trim$(PArray(2)) + ".loc"
    
    If IfExist(FullLocName$) Then
        
        Owner$ = GetLocEntry(LocNo, "3")
        If Owner$ <> UserData(ID, 0) And Val(UserData(ID, 9)) < 2 Then
            SendData ID, "Location" + Str$(LocNo) + " does not belong to you, therefore you cannot change it."
            Exit Sub
        Else
            Call RemoveLocExit(ID, LocNo, DigDir$)
        End If
    
    End If
    Exit Sub
End If

    

    


    
SendData ID, HelpMsg$

End Sub



