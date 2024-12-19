Sub AddToInventory(ID, ItemID)
Dim TmpArray() As String
TmpArray() = Split(UserData(ID, 17), " ")
For t% = 0 To UBound(TmpArray())
    If Val(TmpArray(t%)) <> ItemID Then nArray = nArray + TmpArray(t%) + " "
Next t%
nArray = nArray + Trim$(Str(ItemID))
UserData(ID, 17) = nArray

End Sub

Sub RemoveFromInventory(ID, ItemID)
Dim TmpArray() As String
TmpArray() = Split(UserData(ID, 17), " ")
For t% = 0 To UBound(TmpArray())
If Val(TmpArray(t%)) <> ItemID Then nArray = nArray + TmpArray(t%) + " "
Next t%
UserData(ID, 17) = nArray

End Sub
Sub CheckInventory(ID, ItemID)
CheckInventory = False
Dim TmpArray() As String
TmpArray() = Split(UserData(ID, 17), " ")
For t% = 0 To UBound(TmpArray())
If Val(TmpArray(t%)) <> ItemID Then CheckInventory = True
Next t%
End Sub
