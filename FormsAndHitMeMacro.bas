Attribute VB_Name = "Módulo2"

Public Sub DisplayPlayer1JogarCartas()

Dim form As New Player1JogarCartas

If Range("E27").Value = Range("F6").Value Then
form.Show
End If
End Sub

Public Sub DisplayPlayer2JogarCartas()

Dim form As New Player2JogarCartas

If Range("E27").Value = Range("H6").Value Then
form.Show
End If
End Sub

Public Sub DisplayNovoJogo()

Dim form As New NovoJogo
form.Show

End Sub

Public Sub HitMePlayer1()

Dim lastusedCell As Range
Set lastusedCell = ActiveCell
If Range("E27").Value = Range("F6").Value Then


    Dim mainDeck As Variant
    Dim selectedCard As Integer

    Dim oneCard As Integer
    mainDeck = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)        'main deck

    oneCard = Application.WorksheetFunction.RandBetween(LBound(mainDeck), UBound(mainDeck))   'chooses a random card from the main deck
    selectedCard = mainDeck(oneCard)

    Range("F7:F15").Select
    If Application.WorksheetFunction.CountBlank(Selection) > 0 Then          'if there is room in the table for a card
        Dim Z As Integer
        Z = 0
            For X = 1 To 9
                If Z = 0 Then
                    If IsEmpty(ActiveCell(X, 1)) = True Then                'plays the card in the first empty cell
                        ActiveCell(X, 1).Value = selectedCard
                        Z = Z + 1
                    End If
                End If
                Next X
        lastusedCell.Select

            If Range("F16").Value = 20 Then                               'if with this card, the outcome is 20, pazaak
                Range("D26").Value = "Pazaak"
                    If IsEmpty(Range("F26")) = True Then
                        Range("E27").Value = Range("H6").Value
                    Else
                        Range("E27").Value = "Round Over"
                    End If
            ElseIf Application.WorksheetFunction.CountBlank(Range("F7:F15")) = 0 Then 'if there is no room after playing the card, stand
                    Range("D26").Value = "Stand"
                    If IsEmpty(Range("F26")) = True Then
                    Range("E27").Value = Range("H6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
            Else
                        'if the player still has cards in hand
                If IsEmpty(Range("F19")) = False Or IsEmpty(Range("F20")) = False Or IsEmpty(Range("F21")) = False Or IsEmpty(Range("F22")) = False Then
        
                switchTurn = MsgBox("Do you want to Play a Card(yes), Stand(no), or Continue Playing(cancel)?", vbYesNoCancel, "Decision " & Range("F6").Value)
                    If switchTurn = vbYes Then         'gives him the 3 options
                        Call DisplayPlayer1JogarCartas
                    ElseIf switchTurn = vbNo Then
                        Range("D26").Value = "Stand"
                        If IsEmpty(Range("F26")) = True Then
                            Range("E27").Value = Range("H6").Value
                        Else
                            Range("E27").Value = "Round Over"
                        End If
                    ElseIf switchTurn = vbCancel Then
                        If Range("F16").Value < 20 Then
                            If IsEmpty(Range("F26")) = True Then
                            Range("E27").Value = Range("H6").Value
                            End If
                        Else
                            Range("D26").Value = "Bust"
                            If IsEmpty(Range("F26")) = True Then
                            Range("E27").Value = Range("H6").Value
                            Else
                            Range("E27").Value = "Round Over"
                            End If
                        End If
                    End If
                Else
                If Range("F16").Value > 20 Then
                    Range("D26").Value = "Bust"
                    If IsEmpty(Range("F26")) = True Then
                    Range("E27").Value = Range("H6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
                ElseIf Range("F16").Value = 20 Then
                    Range("D26").Value = "Pazaak"
                    If IsEmpty(Range("F26")) = True Then
                    Range("E27").Value = Range("H6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
                Else
                    switchTurn = MsgBox("Stand(ok) or Continue Playing(cancel)?", vbOKCancel, "Decision " & Range("F6").Value)
                    If switchTurn = vbOK Then      'if the player does not have any cards in hand, shows him only two options, stand or continue
                        Range("D26").Value = "Stand"
                        If IsEmpty(Range("F26")) = True Then
                            Range("E27").Value = Range("H6").Value
                        Else
                            Range("E27").Value = "Round Over"
                        End If
                    ElseIf switchTurn = vbCancel Then
                        If IsEmpty(Range("F26")) = True Then
                            Range("E27").Value = Range("H6").Value
                        End If
                    End If
                End If
        End If
            End If
    Else
            MsgBox "You already have 9 cards in the table, forced Stand", vbOKOnly, "Hit Me " & Range("F6").Value
            lastusedCell.Select
            Range("D26").Value = "Stand"
                If IsEmpty(Range("F26")) = True Then
                    Range("E27").Value = Range("H6").Value
                Else
                    Range("E27").Value = "Round Over"
                End If
    End If
End If

End Sub

Public Sub HitMePlayer2()

Dim lastusedCell As Range
Set lastusedCell = ActiveCell
If Range("E27").Value = Range("H6").Value Then


    Dim mainDeck As Variant
    Dim selectedCard As Integer

    Dim oneCard As Integer
    mainDeck = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

    oneCard = Application.WorksheetFunction.RandBetween(LBound(mainDeck), UBound(mainDeck))
    selectedCard = mainDeck(oneCard)

    Range("H7:H15").Select
    If Application.WorksheetFunction.CountBlank(Selection) > 0 Then
        Dim Z As Integer
        Z = 0
            For X = 1 To 9
                If Z = 0 Then
                    If IsEmpty(ActiveCell(X, 1)) = True Then
                        ActiveCell(X, 1).Value = selectedCard
                        Z = Z + 1
                    End If
                End If
                Next X
        lastusedCell.Select

            If Range("H16").Value = 20 Then
                Range("F26").Value = "Pazaak"
                    If IsEmpty(Range("D26")) = True Then
                        Range("E27").Value = Range("F6").Value
                    Else
                        Range("E27").Value = "Round Over"
                    End If
            ElseIf Application.WorksheetFunction.CountBlank(Range("H7:H15")) = 0 Then
                    Range("F26").Value = "Stand"
                    If IsEmpty(Range("D26")) = True Then
                    Range("E27").Value = Range("F6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
            Else

                If IsEmpty(Range("H19")) = False Or IsEmpty(Range("H20")) = False Or IsEmpty(Range("H21")) = False Or IsEmpty(Range("H22")) = False Then
        
                switchTurn = MsgBox("Do you want to Play a Card(yes), Stand(no), or Continue Playing(cancel)?", vbYesNoCancel, "Decision " & Range("H6").Value)
                    If switchTurn = vbYes Then
                        Call DisplayPlayer2JogarCartas
                    ElseIf switchTurn = vbNo Then
                        Range("F26").Value = "Stand"
                        If IsEmpty(Range("D26")) = True Then
                            Range("E27").Value = Range("F6").Value
                        Else
                            Range("E27").Value = "Round Over"
                        End If
                    ElseIf switchTurn = vbCancel Then
                        If Range("H16").Value < 20 Then
                            If IsEmpty(Range("D26")) = True Then
                            Range("E27").Value = Range("F6").Value
                            End If
                        Else
                            Range("F26").Value = "Bust"
                            If IsEmpty(Range("D26")) = True Then
                            Range("E27").Value = Range("F6").Value
                            Else
                            Range("E27").Value = "Round Over"
                            End If
                        End If
                    End If
                Else
                If Range("H16").Value > 20 Then
                    Range("F26").Value = "Bust"
                    If IsEmpty(Range("D26")) = True Then
                    Range("E27").Value = Range("F6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
                ElseIf Range("H16").Value = 20 Then
                    Range("F26").Value = "Pazaak"
                    If IsEmpty(Range("D26")) = True Then
                    Range("E27").Value = Range("F6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
                Else
                    switchTurn = MsgBox("Stand(ok) or Continue Playing(cancel)?", vbOKCancel, "Decision " & Range("H6").Value)
                    If switchTurn = vbOK Then      ' 2 opções do jogador após o hit, se ele não tem mais carta
                        Range("F26").Value = "Stand"
                        If IsEmpty(Range("D26")) = True Then
                            Range("E27").Value = Range("F6").Value
                        Else
                            Range("E27").Value = "Round Over"
                        End If
                    ElseIf switchTurn = vbCancel Then
                        If IsEmpty(Range("D26")) = True Then
                            Range("E27").Value = Range("F6").Value
                        End If
                    End If
                End If
        End If
            End If
    Else
            MsgBox "You already have 9 cards in the table, forced Stand", vbOKOnly, "Hit Me " & Range("H6").Value
            lastusedCell.Select
            Range("F26").Value = "Stand"
                If IsEmpty(Range("D26")) = True Then
                    Range("E27").Value = Range("F6").Value
                Else
                    Range("E27").Value = "Round Over"
                End If
    End If
End If

End Sub

