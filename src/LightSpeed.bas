Attribute VB_Name = "Module2"
Attribute VB_Name = "Module2"
Sub LightSpeed_Click()
'If death star is up and you have enough hyperfuel to reach lightspeed: you need more than 3 units of fuel to
' enter lightspeed and if so enter the deathstar and change emoji and colour to match the main character
If Selection.Offset(-1, 0).Interior.Color = vbRed And Range("AF7").Value > 2 Then
    Range("AF7").Value = Range("AF7").Value - 3
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(-1, 0).Value = ""
    Selection.Offset(-1, 0).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    MsgBox "You have managed to enter light speed and exit the Death Star unseen!!"
'If death star is up and you don't have enough hyperfuel for lightspeed; teleported back to rebel base or
' the starting point for pressing the button
ElseIf Selection.Offset(-1, 0).Interior.Color = vbRed And Range("AF7") < 3 Then
    MsgBox "I've got a bad feeling about this...."
    MsgBox "Oh NO!! The stormtroopers have caught you but you somehow manage to make a lucky escape and make it back to the Rebel base to recharge!! "
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Range("A17").Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
'If death star is to the right (same logic as before)
ElseIf Selection.Offset(0, 1).Interior.Color = vbRed And Range("AF7").Value > 2 Then
    Range("AF7").Value = Range("AF7").Value - 3
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(0, 1).Value = ""
    Selection.Offset(0, 1).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    MsgBox "You have managed to enter light speed and exit the Death Star unseen!!"
ElseIf Selection.Offset(0, 1).Interior.Color = vbRed And Range("AF7").Value < 3 Then
    MsgBox "Oh NO!! The stormtroopers have caught you but you somehow manage to make a lucky escape and make it back to the Rebel base to recharge!! "
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Range("A17").Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
'If death star is to the left
ElseIf Selection.Offset(0, -1).Interior.Color = vbRed And Range("AF7").Value > 2 Then
    Range("AF7").Value = Range("AF7").Value - 3
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(0, -1).Value = ""
    Selection.Offset(0, -1).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    MsgBox "You have managed to enter light speed and exit the Death Star unseen!!"
ElseIf Selection.Offset(0, -1).Interior.Color = vbRed And Range("AF7").Value < 3 Then
    MsgBox "Oh NO!! The stormtroopers have caught you but you somehow manage to make a lucky escape and make it back to the Rebel base to recharge!! "
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Range("A17").Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
'If death star is below
ElseIf Selection.Offset(1, 0).Interior.Color = vbRed And Range("AF7").Value > 2 Then
    Range("AF7").Value = Range("AF7").Value - 3
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(1, 0).Value = ""
    Selection.Offset(1, 0).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    MsgBox "You have managed to enter light speed and exit the Death Star unseen!!"
ElseIf Selection.Offset(1, 0).Interior.Color = vbRed And Range("AF7").Value < 3 Then
    MsgBox "Oh NO!! The stormtroopers have caught you but you somehow manage to make a lucky escape and make it back to the Rebel base to recharge!! "
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Range("A17").Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
End If
End Sub
