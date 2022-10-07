VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Player2JogarCartas 
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "Player2HandForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Player2JogarCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

If Not Range("H19").Value = "" Then
listCartas.AddItem Range("H19").Value
End If
If Not Range("H20").Value = "" Then
listCartas.AddItem Range("H20").Value
End If
If Not Range("H21").Value = "" Then
listCartas.AddItem Range("H21").Value
End If
If Not Range("H22").Value = "" Then
listCartas.AddItem Range("H22").Value
End If

Dim customCaption As String
customCaption = Range("H6").Value
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

For cardsChosen = 0 To listCartas.ListCount - 1
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

Range("H7:H15").Select
    If IsEmpty(cardArray) = False Then
        If UBound(cardArray) + 1 <= Application.WorksheetFunction.CountBlank(Selection) Then
            Range("H19:H22").Select
            For cardsChosen2 = 0 To listCartas.ListCount - 1
                If listCartas.Selected(cardsChosen2) = True Then
                 cardtobeDeleted = 0
                    For Y = 1 To 4
                    If cardtobeDeleted = 0 Then
                        If InStr(listCartas.List(cardsChosen2), "\") > 0 Then
                            If listCartas.List(cardsChosen2) = ActiveCell(Y, 1).Value Then
                            ActiveCell(Y, 1).ClearContents
                            cardtobeDeleted = cardtobeDeleted + 1
                            End If
                        ElseIf InStr(listCartas.List(cardsChosen2), "&") > 0 Then
                            If listCartas.List(cardsChosen2) = ActiveCell(Y, 1).Value Then
                            ActiveCell(Y, 1).ClearContents
                            cardtobeDeleted = cardtobeDeleted + 1
                            End If
                        Else
                            If InStr(listCartas.List(cardsChosen2), "\") = 0 And InStr(listCartas.List(cardsChosen2), "&") = 0 Then
                                If InStr(ActiveCell(Y, 1), "\") = 0 And InStr(ActiveCell(Y, 1), "&") = 0 Then
                                If Val(listCartas.List(cardsChosen2)) = Val(ActiveCell(Y, 1)) Then
                                ActiveCell(Y, 1).ClearContents
                                cardtobeDeleted = cardtobeDeleted + 1
                                End If
                                End If
                            End If
                        End If
                    End If
                    Next Y
                End If
            Next cardsChosen2

            CardtoPlay = 0
            Range("H7:H15").Select

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
                Range("F26").Value = "Stand"
                If IsEmpty(Range("D26")) = True Then
                    Range("E27").Value = Range("F6").Value
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
    
    If Range("H16").Value = 20 Then
            Range("F26").Value = "Pazaak"
            If IsEmpty(Range("D26")) = True Then
                Range("E27").Value = Range("F6").Value
            Else
                Range("E27").Value = "Round Over"
            End If

    ElseIf Range("H16").Value > 20 Then
                Range("F26").Value = "Bust"
                If IsEmpty(Range("D26")) = True Then
                    Range("E27").Value = Range("F6").Value
                Else
                    Range("E27").Value = "Round Over"
                End If
    Else
    If Application.WorksheetFunction.CountBlank(Range("H7:H15")) = 0 Then
                    If Range("H16").Value > 20 Then
                    Range("F26").Value = "Bust"
                    Else
                    Range("F26").Value = "Stand"
                    End If
                    If IsEmpty(Range("D26")) = True Then
                    Range("E27").Value = Range("F6").Value
                    Else
                    Range("E27").Value = "Round Over"
                    End If
    Else
    
            answer = MsgBox("Stand(ok) or Continue Playing(cancel)?", vbQuestion + vbOKCancel + vbDefaultButton2, "Stand or Continue Playing")
                    If answer = vbCancel Then
                        If IsEmpty(Range("D26")) = True Then
                        Range("E27").Value = Range("F6").Value
                        End If
                    ElseIf answer = vbOK Then
                        Range("F26").Value = "Stand"
                        If IsEmpty(Range("D26")) = True Then
                        Range("E27").Value = Range("F6").Value
                        Else
                        Range("E27").Value = "Round Over"
                        End If
                    End If
     End If
    End If

End Sub


