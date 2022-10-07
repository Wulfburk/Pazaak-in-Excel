VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Player1JogarCartas 
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "Player1HandForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Player1JogarCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

If Not Range("F19").Value = "" Then
listCartas.AddItem Range("F19").Value
End If
If Not Range("F20").Value = "" Then
listCartas.AddItem Range("F20").Value
End If
If Not Range("F21").Value = "" Then
listCartas.AddItem Range("F21").Value
End If
If Not Range("F22").Value = "" Then
listCartas.AddItem Range("F22").Value
End If

Dim customCaption As String
customCaption = Range("F6").Value       'adds the player name to the panel
Label2.Caption = customCaption & "'s Hand"


End Sub

Private Sub Jogar_Click()

Dim lastusedCell As Range
Set lastusedCell = ActiveCell


    Dim cardsChosen As Integer
    Dim Count As Integer

    Dim TableCards As Integer
    Dim CardtoPlay As Integer
    Dim cardArray As Variant


    Count = 0

    For cardsChosen = 0 To listCartas.ListCount - 1                         'if a card in the list was chosen, then adds that card to the array
        If listCartas.Selected(cardsChosen) = True Then
            If Count = 0 Then
                ReDim cardArray(Count)
            Else
                ReDim Preserve cardArray(Count)
            End If
            cardArray(Count) = listCartas.List(cardsChosen)
            Count = Count + 1
        End If
    Next cardsChosen

    Range("F7:F15").Select
        If IsEmpty(cardArray) = False Then
            If UBound(cardArray) + 1 <= Application.WorksheetFunction.CountBlank(Selection) Then
            Range("F19:F22").Select                                                          'checks if there is room in the table for the selected cards
            For cardsChosen2 = 0 To listCartas.ListCount - 1                                   'plays the cards and deletes them from the hand
                If listCartas.Selected(cardsChosen2) = True Then
                cardtobeDeleted = 0
                    For Y = 1 To 4
                    If cardtobeDeleted = 0 Then
                        If InStr(listCartas.List(cardsChosen2), "\") > 0 Then 'if the card has a "\" then sees that it is special
                            If listCartas.List(cardsChosen2) = ActiveCell(Y, 1).Value Then
                            ActiveCell(Y, 1).ClearContents                     'deletes the card with "\" with equal numbers to the selected
                            cardtobeDeleted = cardtobeDeleted + 1 'makes sure that only one card equal to the selected, is deleted
                            End If
                        ElseIf InStr(listCartas.List(cardsChosen2), "&") > 0 Then 'if the card has a "&", sees that it is special
                            If listCartas.List(cardsChosen2) = ActiveCell(Y, 1).Value Then
                            ActiveCell(Y, 1).ClearContents                   'deletes the card with & with equal numbers as the selected
                            cardtobeDeleted = cardtobeDeleted + 1
                            End If
                        Else
                        If InStr(listCartas.List(cardsChosen2), "\") = 0 And InStr(listCartas.List(cardsChosen2), "&") = 0 Then
                            If InStr(ActiveCell(Y, 1), "\") = 0 And InStr(ActiveCell(Y, 1), "&") = 0 Then
                            If Val(listCartas.List(cardsChosen2)) = Val(ActiveCell(Y, 1)) Then 'if the card has no \ our &
                            ActiveCell(Y, 1).ClearContents                              'deletes the card with equal numbers as the selected
                            cardtobeDeleted = cardtobeDeleted + 1                        'makes sure that it is deleted only once
                            End If
                            End If
                        End If
                        End If
                    End If
                    Next Y
                End If
            Next cardsChosen2

            CardtoPlay = 0
            Range("F7:F15").Select

            For TableCards = 1 To 9
                If IsEmpty(ActiveCell(TableCards, 1)) = True Then
                    ActiveCell(TableCards, 1).Value = cardArray(CardtoPlay)
                    CardtoPlay = CardtoPlay + 1
                End If
                If CardtoPlay > UBound(cardArray) Then
                    Exit For
                End If
            Next TableCards
    
                If Application.WorksheetFunction.CountBlank(Range("H7:H15")) = 0 Then
                    Range("D26").Value = "Stand"
                    If IsEmpty(Range("F26")) = True Then
                        Range("E27").Value = Range("H6").Value
                    Else
                        Range("E27").Value = "Round Over"
                    End If
                End If
        lastusedCell.Select
        Unload Me
        Else
            lastusedCell.Select
            MsgBox "You cannot have more than 9 cards in the table. Select fewer cards, or play none", vbOKOnly, "Play Cards"
        End If
Else
    lastusedCell.Select
    Unload Me
End If
            If Range("F16").Value = 20 Then
                    Range("D26").Value = "Pazaak"
                    If IsEmpty(Range("F26")) = True Then
                        Range("E27").Value = Range("H6").Value
                    Else
                        Range("E27").Value = "Round Over"
                    End If
            ElseIf Range("F16").Value > 20 Then
                Range("D26").Value = "Bust"
                If IsEmpty(Range("F26")) = True Then
                Range("E27").Value = Range("H6").Value
                Else
                Range("E27").Value = "Round Over"
                End If
            Else
            
            If Application.WorksheetFunction.CountBlank(Range("F7:F15")) = 0 Then 'If with the played cards, the number of cards in the table has reached 9, forces stand or bust.
                    If Range("F16").Value > 20 Then
                    Range("D26").Value = "Bust"
                    Else
                    Range("D26").Value = "Stand"
                    End If
                    If IsEmpty(Range("F26")) = True Then
                    Range("E27").Value = Range("H6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
            Else
            
                   'otherwise, allows to continue playing or stand
            answer = MsgBox("Stand(ok) or Continue Playing(cancel)?", vbQuestion + vbOKCancel + vbDefaultButton2, "Stand or Continue Playing")
                If answer = vbCancel Then                           'switches turn if the other player is NOT stand or bust, otherwise, round over
                        If IsEmpty(Range("F26")) = True Then
                        Range("E27").Value = Range("H6").Value
                        End If
                ElseIf answer = vbOK Then
                    Range("D26").Value = "Stand"
                    If IsEmpty(Range("F26")) = True Then
                    Range("E27").Value = Range("H6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
                End If
            End If
            End If

End Sub

