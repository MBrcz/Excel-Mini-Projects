Attribute VB_Name = "TTTLogic"
' This module contains functions related to the win conditions and finishing game.
Option Explicit

Public Function IsAnyMoveValid(board_table As Variant) As Boolean
' Checks if any move can be performed by any player.

    Dim num As Variant

    For Each num In board_table
        If num = eBoardMove.Unoccupied Then
            IsAnyMoveValid = True
            Exit Function
        End If
    Next num

    IsAnyMoveValid = False

End Function

Public Function HasPlayerWin(player_number As Byte, board_table As Variant) As Boolean
' Checks if player has won the game.

    Dim winPatterns(7, 2) As Integer
    Dim i As Integer

    ' All possible winning patterns:
    winPatterns(0, 0) = 0: winPatterns(0, 1) = 1: winPatterns(0, 2) = 2
    winPatterns(1, 0) = 3: winPatterns(1, 1) = 4: winPatterns(1, 2) = 5
    winPatterns(2, 0) = 6: winPatterns(2, 1) = 7: winPatterns(2, 2) = 8
    winPatterns(3, 0) = 0: winPatterns(3, 1) = 3: winPatterns(3, 2) = 6
    winPatterns(4, 0) = 1: winPatterns(4, 1) = 4: winPatterns(4, 2) = 7
    winPatterns(5, 0) = 2: winPatterns(5, 1) = 5: winPatterns(5, 2) = 8
    winPatterns(6, 0) = 0: winPatterns(6, 1) = 4: winPatterns(6, 2) = 8
    winPatterns(7, 0) = 2: winPatterns(7, 1) = 4: winPatterns(7, 2) = 6

    ' Check if any winning pattern matches the board_table
    For i = 0 To UBound(winPatterns, 1)
        If board_table(winPatterns(i, 0)) = player_number _
        And board_table(winPatterns(i, 1)) = player_number _
        And board_table(winPatterns(i, 2)) = player_number Then
            HasPlayerWin = True
            Exit Function
        End If
    Next i

    HasPlayerWin = False

End Function

Public Function CountNonEmptyBoardSpaces(board_table As Variant) As Byte
' Checks how many there are not empty places in the board.
    
    Dim count As Variant
    Dim num As Byte
    
    For Each count In board_table
        If count <> eBoardMove.Unoccupied Then
            num = num + 1
        End If
    Next count

    CountNonEmptyBoardSpaces = num
    
End Function

Public Function FinishGame(player As clsPlayer, does_player_win As Boolean)
' Function that handles exiting the game.
    
    Dim user_question As String
    Dim prompt As String
    
    If does_player_win = True Then
        player.LetWonGames = player.GetWonGames() + 1
        
        prompt = "Game finished! " & vbCrLf & "The player named " & player.GetName & " has won!"
    
    ElseIf does_player_win = False Then
        prompt = "The game has finished in stalemate!"
    
    End If
    
    user_question = MsgBox(prompt & vbCrLf & "Do you want to play again?", vbYesNo + vbQuestion, "Question")
    
    If user_question = vbYes Then
       Main.RestartGame
      
    ElseIf user_question = vbNo Then
       Main.ExitGame

    End If
        
End Function

