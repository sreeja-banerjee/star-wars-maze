Attribute VB_Name = "Module1"
Attribute VB_Name = "Module1"
Sub Up_Click()
' Take the fuel from the box, change colour and add to fuel count: remove emoji and colour of original
' and add to the location of the fuel while removing the colour and fuel emoji, each emoji represents 1 unit
'of fuel which will be added to the fuel counter
If Selection.Offset(-1, 0).Interior.Color = vbCyan Then
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(-1, 0).Value = ""
    Selection.Offset(-1, 0).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    Range("AF7").Value = Range("AF7").Value + 1
'Out of range from the maze: to prevent the player from going into the black boxes or out of the boundaries of the maze
ElseIf Selection.Offset(-1, 0).Interior.Color = vbBlack Or Selection.Offset(-1, 0).Interior.Color = vbWhite Then
    MsgBox "You are venturing into deep space! Danger!"
'Instructions for presence of Death Star: call for action for the player to use the special lightspeed button
ElseIf Selection.Offset(-1, 0).Interior.Color = vbRed Then
    MsgBox "Danger, First Order's presence has been detected, use LightSpeed to evade the Death Star!!"
'For receiving advice from the wise descendent of Yoga and adding to Advice counter: The advice received from
' the NPC which will be taken into account for the end of the game by recording it at the advice counter
ElseIf Selection.Offset(-1, 0).Interior.Color = vbBlue Then
    MsgBox "You have entered the swamp planet of Dagobah"
    MsgBox "You find the descendent of Master Yoda and he tells you 'Do or do not. There is no Try'.."
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(-1, 0).Value = ""
    Selection.Offset(-1, 0).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    Range("AF9").Value = "Yes"
' Going to the princess without receiving advice: you will not be able to finish the game without the presence
' of the advice from NPC therefore call for action given to go and receive it
ElseIf Selection.Offset(-1, 0).Interior.Color = vbMagenta And Range("AF9").Value = "" Then
    MsgBox "You are not yet wise enough for the princess!"
' Going to the princess with the advice aka game over: you have received the advice and reached the end of the
' maze, changing emojis to represent 2 characters together with the game over message
ElseIf Selection.Offset(-1, 0).Interior.Color = vbMagenta And Range("AF9").Value = "Yes" Then
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(-1, 0).Select
    Selection.Interior.Color = vbWhite
    Selection.Value = ChrW(55357) & ChrW(56427)
    MsgBox "And they lived happily ever after till the end of their reign. Game Over!!"
' Travelling on all other possible spots on the maze
Else
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(-1, 0).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
End If

End Sub
Sub StartGame_Click()
' main character
Range("A17").Select
Selection.Value = ChrW(55357) & ChrW(56424)
' yoda descendent
Range("T4").Value = ChrW(55357) & ChrW(56445)
' death stars
Range("E9").Value = ChrW(55356) & ChrW(57105)
Range("G6").Value = ChrW(55356) & ChrW(57105)
Range("L5").Value = ChrW(55356) & ChrW(57105)
Range("R9").Value = ChrW(55356) & ChrW(57105)
Range("Y5").Value = ChrW(55356) & ChrW(57105)
' princess
Range("AA2").Value = ChrW(55357) & ChrW(56440)
' hyperfuel location
Range("C13").Value = ChrW(55357) & ChrW(56485)
Range("E15").Value = ChrW(55357) & ChrW(56485)
Range("G17").Value = ChrW(55357) & ChrW(56485)
Range("M17").Value = ChrW(55357) & ChrW(56485)
Range("T15").Value = ChrW(55357) & ChrW(56485)
Range("Y16").Value = ChrW(55357) & ChrW(56485)
Range("H13").Value = ChrW(55357) & ChrW(56485)
Range("O13").Value = ChrW(55357) & ChrW(56485)
Range("Q12").Value = ChrW(55357) & ChrW(56485)
Range("K10").Value = ChrW(55357) & ChrW(56485)
Range("C9").Value = ChrW(55357) & ChrW(56485)
Range("V9").Value = ChrW(55357) & ChrW(56485)
Range("Z9").Value = ChrW(55357) & ChrW(56485)
Range("L8").Value = ChrW(55357) & ChrW(56485)
Range("V7").Value = ChrW(55357) & ChrW(56485)
Range("S7").Value = ChrW(55357) & ChrW(56485)
Range("K6").Value = ChrW(55357) & ChrW(56485)
Range("C6").Value = ChrW(55357) & ChrW(56485)
Range("E3").Value = ChrW(55357) & ChrW(56485)
Range("V3").Value = ChrW(55357) & ChrW(56485)
Range("N2").Value = ChrW(55357) & ChrW(56485)
' fuel count
Range("AF7") = "0"
' Advice counter
Range("AF9") = ""


End Sub
Sub Down_Click()
' Take the fuel from the box, change colour and add to fuel count
If Selection.Offset(1, 0).Interior.Color = vbCyan Then
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(1, 0).Value = ""
    Selection.Offset(1, 0).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    Range("AF7").Value = Range("AF7").Value + 1
'Out of range from the maze
ElseIf Selection.Offset(1, 0).Interior.Color = vbBlack Or Selection.Offset(1, 0).Interior.Color = vbWhite Then
        MsgBox "You are venturing into deep space! Danger!"
'Death Star Instructions
ElseIf Selection.Offset(1, 0).Interior.Color = vbRed Then
    MsgBox "Danger, First Order's presence has been detected, use LightSpeed to evade the Death Star!!"
' Travelling on all other possible spots on the maze
Else
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(1, 0).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
End If

End Sub
Sub Right_Click()
' Take the fuel from the box, change colour and add to fuel count
If Selection.Offset(0, 1).Interior.Color = vbCyan Then
   Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(0, 1).Value = ""
    Selection.Offset(0, 1).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    Range("AF7").Value = Range("AF7").Value + 1
'Out of range from maze
ElseIf Selection.Offset(0, 1).Interior.Color = vbBlack Or Selection.Offset(0, 1).Interior.Color = vbWhite Then
        MsgBox "You are venturing into deep space! Danger!"
'Intructions for Death Star
ElseIf Selection.Offset(0, 1).Interior.Color = vbRed Then
    MsgBox "Danger, First Order's presence has been detected, use LightSpeed to evade the Death Star!!"
' Travelling on all other possible spots on the maze
Else
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(0, 1).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
End If

End Sub

Sub Reset_Click()
Attribute Reset_Click.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Reset_Click Macro
'

'
    Sheets("Copy").Select
    Range("A1:AA17").Select
    Selection.Copy
    Sheets("Challenge 1").Select
    Range("A1").Select
    ActiveSheet.Paste
End Sub
Sub Left_Click()
' Take the fuel from the box, change colour and add to fuel count
If Selection.Offset(0, -1).Interior.Color = vbCyan Then
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(0, -1).Value = ""
    Selection.Offset(0, -1).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
    Range("AF7").Value = Range("AF7").Value + 1
'Out of range from the maze
ElseIf Selection.Offset(0, -1).Interior.Color = vbBlack Or Selection.Offset(0, -1).Interior.Color = vbWhite Then
        MsgBox "You are venturing into deep space! Danger!"
'Death Star Instructions
ElseIf Selection.Offset(0, -1).Interior.Color = vbRed Then
    MsgBox "Danger, First Order's presence has been detected, use LightSpeed to evade the Death Star!!"
' Travelling on all other possible spots on the maze
Else
    Selection.Interior.Color = vbYellow
    Selection.Value = ""
    Selection.Offset(0, -1).Select
    Selection.Interior.Color = vbGreen
    Selection.Value = ChrW(55357) & ChrW(56424)
End If

End Sub
