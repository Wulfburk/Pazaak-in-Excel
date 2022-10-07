VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NovoJogo 
   Caption         =   "Novo Jogo?"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "NewMatchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NovoJogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

Player1Name.Text = Range("F6").Value
Player2Name.Text = Range("H6").Value

End Sub


Private Sub Reset_Click()
answer = MsgBox("Are you sure you want to reset scores, player names and clear the table?", vbQuestion + vbYesNo + vbDefaultButton2, "Reset")
If answer = vbYes Then
Range("K28:L28").Value = 0
Range("H27:H31").ClearContents           'clears scores and the table
Range("F7:F15").ClearContents
Range("H7:H15").ClearContents
Range("D26").ClearContents
Range("F26").ClearContents
Range("E27").ClearContents
Range("F19:F22").ClearContents
Range("H19:H22").ClearContents

Range("F6").Value = "Player 1"           'returns player names to default
Range("H6").Value = "Player 2"
Player1Name.Text = Range("F6").Value
Player2Name.Text = Range("H6").Value

End If
End Sub

Private Sub NewGame_Click()

Dim lastusedCell As Range
Set lastusedCell = ActiveCell

answer = MsgBox("are you sure you want to start a new match?", vbQuestion + vbYesNo + vbDefaultButton2, "New Match")
If answer = vbYes Then

Range("F6").Value = Player1Name.Text   'Gets the player names from the text box
Range("H6").Value = Player2Name.Text   'and allows them to be referenced throughout the game


Dim startingPlayer As Integer
startingPlayer = Application.WorksheetFunction.RandBetween(1, 2)
Range("K27:L27").Select
Range("E27").Value = ActiveCell(1, startingPlayer).Value 'randomizes who will start the match

Dim playerDeck As Variant
Dim selectedCards1(0 To 3) As Variant
Dim selectedCards2(0 To 3) As Variant

Dim X As Integer
Dim oneCard As Integer
playerDeck = Array(1, 2, 3, 4, 5, 6, -1, -2, -3, -4, -5, -6, "1 \ -1", "2 \ -2", "3 \ -3", "4 \ -4", "5 \ -5", "6 \ -6", "2 & 4", "3 & 6")
                'establishes the player deck
For X = 0 To 3
oneCard = Application.WorksheetFunction.RandBetween(LBound(playerDeck), UBound(playerDeck))
selectedCards1(X) = playerDeck(oneCard) 'chooses 4 random cards from the deck
Next X
Range("F19:F22").Select
For X = 1 To 4
ActiveCell(X, 1).Value = selectedCards1(X - 1) 'places the 4 cards in the players' hands
Next X

For X = 0 To 3
oneCard = Application.WorksheetFunction.RandBetween(LBound(playerDeck), UBound(playerDeck))
selectedCards2(X) = playerDeck(oneCard)
Next X
Range("H19:H22").Select
For X = 1 To 4
ActiveCell(X, 1).Value = selectedCards2(X - 1)
Next X

Range("H27:H31").ClearContents
Range("F7:F15").ClearContents
Range("H7:H15").ClearContents
Range("D26").ClearContents
Range("F26").ClearContents

Dim quoteArray As Variant
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
Range("C3").Value = quoteArray(quoteNumber) 'chooses a new quote to appear at the top of the table

lastusedCell.Select
Unload Me
End If
End Sub

Private Sub Sair_Click()
Unload Me
End Sub
