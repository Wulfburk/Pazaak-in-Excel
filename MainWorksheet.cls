VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


Private Sub Worksheet_Change(ByVal Target As Range)

'if E27 shows Round Over, grabs the winner in K16
'and places him at the Round Wins range
'Clears the content of the table, but not the players' hands and scores
'Then chooses who will start the next round

If Target.Address = "$E$27" Then
Dim lastusedCell As Range
Set lastusedCell = ActiveCell
Dim Z As Integer
Dim W As Integer
Dim Y As Integer
If Range("E27").Value = "Round Over" Then

Dim quoteArray As Variant 'array of quotes to appear at the top of the table
quoteArray = Array("Pure Pazaak! - Atton Rand", "How about a game of Pazaak, Republic Senate rules? - Meetra Surik", _
"Pazaak bores me. I often suspect my opponent of cheating. I prefer predictable games, such as galactic economics - GOTO", _
"Time to even the odds! - Atton Rand", "This better not be using Nar Shaddaa Rules - Meetra Surik", _
"Nar Shaddaa may be one of the biggest cesspits in the galaxy, but it's got a life to it, activity. Aliens, people, refugees... - Mira", _
"[Nar Shaddaa's] like noise, but relaxing. Like the hum of a hyperdrive - Mira", _
"Mucha Shaka Paka - A Twilek", "I have had enough of this. I will be in my chambers - Kreia", _
"Who let the likes of you in here? - A Quarren", _
"I have time for Pazaak, Human, no time for anything else - Geredi", _
"You are big stuff, no? - The Champ", "Fiu fiu fiu fiu - The Champ", _
"Never have I been to a place so alive with the Force, yet so dead to it. The contrast is like a blade - Visas Marr", _
"Ah, the beautiful stench of decay and desperate living - Atton Rand", _
"Welcome to Nar Shaddaa: towering buildings kilometers high and miles deep. Watch where you step, or you'll fall for hours - Atton Rand", _
"Dwooooo. Deet deet deet - Bao-Dur's Remote")

quoteNumber = Application.WorksheetFunction.RandBetween(LBound(quoteArray), UBound(quoteArray))
Range("C3").Value = quoteArray(quoteNumber) 'chooses a new random quote to appear at the top of the table


    If Range("K16").Value = Range("F6").Value Then
        MsgBox "Round is over and the winner is " & Range("F6").Value, vbOKOnly, "Round Winner"
        Range("H27:H31").Select
        Z = 0
        For W = 1 To 5
            If Z = 0 Then
                If IsEmpty(ActiveCell(W, 1)) = True Then
                ActiveCell(W, 1).Value = Range("F6").Value
                Z = Z + 1
                End If
            End If
        Next W
        Range("F7:F15,H7:H15").ClearContents 'clears the table
        Range("D26,F26").ClearContents           'clears the status of the players, like bust or stand
        Range("E27").Value = Range("F6").Value   'If player 1 won, he stars the next round
        lastusedCell.Select
    ElseIf Range("K16").Value = Range("H6").Value Then
        MsgBox "Round is over and the winner is " & Range("H6").Value, vbOKOnly, "Round Winner"
        Range("H27:H31").Select
        Z = 0
        For W = 1 To 5
        If Z = 0 Then
        If IsEmpty(ActiveCell(W, 1)) = True Then
        ActiveCell(W, 1).Value = Range("H6").Value
        Z = Z + 1
        End If
        End If
        Next W
        Range("F7:F15,H7:H15").ClearContents
        Range("D26,F26").ClearContents
        Range("E27").Value = Range("H6").Value   'If player 2 won, he starts the next round
        lastusedCell.Select
    Else
        MsgBox "Round was a draw!", vbOKOnly, "Draw"
        Range("F7:F15,H7:H15").ClearContents
        Range("D26,F26").ClearContents
    
        Range("H27:H31").Select
        Z = 0                           'if it was a draw and one side already won a round
        For W = 5 To 1 Step -1          'the last winner starts the next round
         If Z = 0 Then
            If IsEmpty(ActiveCell(W, 1)) = False Then
                If ActiveCell(W, 1).Value = Range("F6").Value Then
                Range("E27").Value = Range("H6").Value
                MsgBox "Round was a draw. As " & Range("F6").Value & " was the last round winner, " & Range("H6").Value & " starts the next round", vbOKOnly, "Draw"
                ElseIf ActiveCell(W, 1).Value = Range("H6").Value Then
                Range("E27").Value = Range("F6").Value
                MsgBox "Round was a draw. As " & Range("H6").Value & " was the last round winner, " & Range("F6").Value & " starts the next round", vbOKOnly, "Draw"
                End If
                Z = Z + 1
             End If
         End If
        Next W
        Z = 0
        For Y = 1 To 5
            If IsEmpty(ActiveCell(Y, 1)) = True Then
            Z = Z + 1
            End If
        Next Y 'if it was a draw and no side won yet, chooses randomly who will start net round
        If Z = 5 Then
            startingPlayer = Application.WorksheetFunction.RandBetween(1, 2)
            Range("K27:L27").Select
            Range("E27").Value = ActiveCell(1, startingPlayer).Value 'randomiza quem vai começar a partida
            MsgBox "Round was a draw. As there was no previous round winner, who starts the next round was chosen randomly", vbOKOnly, "Draw"
        End If
        lastusedCell.Select
    End If
End If
End If

If Target.Address = "$F$7" Or Target.Address = "$F$8" Or Target.Address = "$F$9" _
Or Target.Address = "$F$10" Or Target.Address = "$F$11" Or Target.Address = "$F$12" _
Or Target.Address = "$F$13" Or Target.Address = "$F$14" Or Target.Address = "$F$15" Then

'Checks if a special card was used in the table
'and if it was, opens the message box for the player to decide if the card is Positive Or Negative
'changes the value of the card in the table, according to the player's decision

Set lastusedCell = ActiveCell
Range("F7:F15").Select
Dim playercardsSpecial As Integer
For playercardsSpecial = 1 To 9
    Select Case ActiveCell(playercardsSpecial, 1).Value
    
    Case "1 \ -1"
    answer = MsgBox("Do you want to play 1 \ -1 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 1
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -1
    End If
    
    Case "2 \ -2"
    answer = MsgBox("Do you want to play 2 \ -2 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 2
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -2
    End If
    
    Case "3 \ -3"
    answer = MsgBox("Do you want to play 3 \ -3 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 3
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -3
    End If

    Case "4 \ -4"
    answer = MsgBox("Do you want to play 4 \ -4 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 4
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -4
    End If
    
    Case "5 \ -5"
    answer = MsgBox("Do you want to play 5 \ -5 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 5
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -5
    End If
    
    Case "6 \ -6"
    answer = MsgBox("Do you want to play 6 \ -6 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 6
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -6
    End If
    
    Case "2 & 4"
    answer = MsgBox("2 & 4 card inverts the value of the 3 or 6 cards that are already in your table", vbOKOnly, "Special Card")
    If answer = vbOK Then
    ActiveCell(playercardsSpecial, 1).Value = "2 @ 4"
        For cardtobeInverted = 1 To 9
        If InStr(ActiveCell(cardtobeInverted, 1), "@") = 0 Then
        If InStr(ActiveCell(cardtobeInverted, 1), "2") > 0 Or InStr(ActiveCell(cardtobeInverted, 1), "4") > 0 Then
        cellvalue = ActiveCell(cardtobeInverted, 1).Value
        ActiveCell(cardtobeInverted, 1).Value = cellvalue * -1
        End If
        End If
        Next cardtobeInverted
    End If
    
    Case "3 & 6"
    answer = MsgBox("3 & 6 card inverts the value of the 3 or 6 cards that are already in your table", vbOKOnly, "Special Card")
    If answer = vbOK Then
    ActiveCell(playercardsSpecial, 1).Value = "3 @ 6"
        For cardtobeInverted = 1 To 9
        If InStr(ActiveCell(cardtobeInverted, 1), "@") = 0 Then
        If InStr(ActiveCell(cardtobeInverted, 1), "3") > 0 Or InStr(ActiveCell(cardtobeInverted, 1), "6") > 0 Then
        cellvalue = ActiveCell(cardtobeInverted, 1).Value
        ActiveCell(cardtobeInverted, 1).Value = cellvalue * -1
        End If
        End If
        Next cardtobeInverted
    End If
    
    
    
    
    End Select
Next playercardsSpecial
lastusedCell.Select
End If

If Target.Address = "$H$7" Or Target.Address = "$H$8" Or Target.Address = "$H$9" _
Or Target.Address = "$H$10" Or Target.Address = "$H$11" Or Target.Address = "$H$12" _
Or Target.Address = "$H$13" Or Target.Address = "$H$14" Or Target.Address = "$H$15" Then


Set lastusedCell = ActiveCell
Range("H7:H15").Select
For playercardsSpecial = 1 To 9
    Select Case ActiveCell(playercardsSpecial, 1).Value
    
    Case "1 \ -1"
    answer = MsgBox("Do you want to play 1 \ -1 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 1
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -1
    End If
    
    Case "2 \ -2"
    answer = MsgBox("Do you want to play 2 \ -2 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 2
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -2
    End If
    
    Case "3 \ -3"
    answer = MsgBox("Do you want to play 3 \ -3 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 3
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -3
    End If

    Case "4 \ -4"
    answer = MsgBox("Do you want to play 4 \ -4 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 4
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -4
    End If
    
    Case "5 \ -5"
    answer = MsgBox("Do you want to play 5 \ -5 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 5
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -5
    End If
    
    Case "6 \ -6"
    answer = MsgBox("Do you want to play 6 \ -6 as a positive(yes) or negative(no)?", vbYesNo, "Special Card")
    If answer = vbYes Then
    ActiveCell(playercardsSpecial, 1).Value = 6
    ElseIf answer = vbNo Then
    ActiveCell(playercardsSpecial, 1).Value = -6
    End If
    
    Case "2 & 4"
    answer = MsgBox("2 & 4 card inverts the value of the 2 or 4 cards that are already in your table", vbOKOnly, "Special Card")
    If answer = vbOK Then
    ActiveCell(playercardsSpecial, 1).Value = "2 @ 4"
        For cardtobeInverted2 = 1 To 9
        If InStr(ActiveCell(cardtobeInverted2, 1), "@") = 0 Then
        If InStr(ActiveCell(cardtobeInverted2, 1), "2") > 0 Or InStr(ActiveCell(cardtobeInverted2, 1), "4") > 0 Then
        cellvalue = ActiveCell(cardtobeInverted2, 1).Value
        ActiveCell(cardtobeInverted2, 1).Value = cellvalue * -1
        End If
        End If
        Next cardtobeInverted2
    End If
    
    Case "3 & 6"
    answer = MsgBox("3 & 6 card inverts the value of the 3 or 6 cards that are already in your table", vbOKOnly, "Special Card")
    If answer = vbOK Then
    ActiveCell(playercardsSpecial, 1).Value = "3 @ 6"
        For cardtobeInverted2 = 1 To 9
        If InStr(ActiveCell(cardtobeInverted2, 1), "@") = 0 Then
        If InStr(ActiveCell(cardtobeInverted2, 1), "3") > 0 Or InStr(ActiveCell(cardtobeInverted2, 1), "6") > 0 Then
        cellvalue2 = ActiveCell(cardtobeInverted2, 1).Value
        ActiveCell(cardtobeInverted2, 1).Value = cellvalue2 * -1
        End If
        End If
        Next cardtobeInverted2
    End If
    
    
    End Select
Next playercardsSpecial
lastusedCell.Select
End If


End Sub

Private Sub Worksheet_Calculate()

If Range("I29").Value = Range("F6").Value Then 'If there was a winner, clears the table, shows a message with the winner, and adds to his score
    MsgBox "Congratulations! " & Range("F6").Value & " won the match due to winning three rounds!", vbOKOnly, "Match is Over"
    Range("H27:H31").ClearContents         'Clears round results
    Range("F19:F22").ClearContents 'clears player 1 hand
    Range("H19:H22").ClearContents 'clears player 2 hand
    Range("F7:F15").ClearContents 'clears player 1 table
    Range("H7:H15").ClearContents 'clears player 2 table
    Range("D26").ClearContents              'clears player 1 status
    Range("F26").ClearContents              'clears player 2 status
    Range("E27").ClearContents              'clears who will begin the match
    Range("K28").Value = Val(Range("K28")) + 1 'adds +1 to player 1 score
ElseIf Range("I29").Text = Range("H6").Value Then
    MsgBox "Congratulations! " & Range("h6").Value & " won the match due to winning three rounds!", vbOKOnly, "Match is Over"
    Range("H27:H31").ClearContents
    Range("F19:F22").ClearContents
    Range("H19:H22").ClearContents
    Range("F7:F15").ClearContents
    Range("H7:H15").ClearContents
    Range("D26").ClearContents
    Range("F26").ClearContents
    Range("E27").ClearContents
    Range("L28").Value = Val(Range("L28")) + 1
End If
End Sub


