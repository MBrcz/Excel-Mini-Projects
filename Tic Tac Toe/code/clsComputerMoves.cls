VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComputerMoves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private CurrentPlayerNum As Byte
Private OppositePlayerNum As Byte

Public Property Let LetCurrentPlayerNum(value As Byte)
    
    CurrentPlayerNum = value

End Property

Public Property Let LetOppositePlayerNum(value As Byte)

    OppositePlayerNum = value

End Property

Private Property Get GetCurrentPlayerNum()

    GetCurrentPlayerNum = CurrentPlayerNum
    
End Property

Private Property Get GetOppositePlayerNum()

    GetOppositePlayerNum = OppositePlayerNum
    
End Property

Public Function MakeMove(mode As String, board_table As Variant)
' Chooses appropriate move that depends on the mode selected in settings.
    
    Dim chosen_move As Byte

    Select Case mode
        
        Case Enums.TranslateComputerMode(UNBEATABLE)
            chosen_move = ChooseOptimalMove(board_table)
        
        Case Enums.TranslateComputerMode(random)
            chosen_move = ChooseRandomMove(board_table)
        
        Case Enums.TranslateComputerMode(mixed)
            chosen_move = ChooseMixedMove(board_table)
    
    End Select

    MakeMove = chosen_move
    
End Function

' ----------------------------------------------------------------------------------------------------------------
' -------------------------------------- MIXED MODE -----------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------

Private Function ChooseMixedMove(board_table) As Byte
' Choses optimal or random move - depends on pseudo RNG mechanism
    
    Dim random_chance As Integer
    
    random_chance = Application.WorksheetFunction.RandBetween(0, 100)
    
    If random_chance >= 50 Then
    
        Debug.Print "Chance that was set: " & CStr(random_chance) & " OPTIMAL mode has been chosen"
        ChooseMixedMove = ChooseOptimalMove(board_table)
    ElseIf random_chance < 50 Then
    
        Debug.Print "Chance that was set: " & CStr(random_chance) & " RANDOM mode has been chosen"
        ChooseMixedMove = ChooseRandomMove(board_table)
    
    End If
    
End Function

' ----------------------------------------------------------------------------------------------------------------
' ---------------------------------- RANDOM MODE -----------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------

Private Function ChooseRandomMove(board_table) As Byte
' Chooses completely random move.

    Dim valid_moves() As Variant
    Dim count As Byte
    Dim move As Byte
    
    count = 0
    For move = LBound(board_table) To UBound(board_table)
        If board_table(move) = 0 Then
            ReDim Preserve valid_moves(count)
            valid_moves(count) = move
            count = count + 1
        End If
    Next move

    Dim random_index As Integer
    random_index = utils.GetRandomIndexFromArray(valid_moves)
    ChooseRandomMove = valid_moves(random_index)
    
End Function

' ----------------------------------------------------------------------------------------------------------------
' ---------------------------------- OPTIMAL MODE ------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------

' BASED ON: https://saturncloud.io/blog/minimax-algorithm-for-tic-tac-toe-in-python/

Private Function ChooseOptimalMove(board_table) As Byte
' Chooses the best move according to minmax Alghoritm.
    
    Dim optimal_score As Double
    Dim local_score As Double

    Dim move As Byte
    Dim optimal_move As Byte
    
    ' If middle point is free, then get it.
    If board_table(eBoardMove.MiddlePoint) = eBoardMove.Unoccupied Then
        ChooseOptimalMove = eBoardMove.MiddlePoint
        Exit Function
    End If
    
    optimal_score = eMinMaxScores.MAXIMUMCOMPUTER
    Const depth = 0
    
     For move = LBound(board_table) To UBound(board_table)
        If board_table(move) = eBoardMove.Unoccupied Then
            
            board_table(move) = GetCurrentPlayerNum()
            local_score = MinMaxAlghoritm(board_table, depth, False)
            board_table(move) = eBoardMove.Unoccupied
            
            If local_score > optimal_score Then
                optimal_move = move
                optimal_score = local_score
              
            End If
            
        End If
    Next move
    
    ChooseOptimalMove = optimal_move

End Function

Private Function MinMaxAlghoritm(board_table As Variant, depth As Byte, is_maximizer_turn As Boolean) As Double
' Recursive function that gets the best score.

    Dim score As Double
    Dim maximum_eval As Double
    Dim minimum_eval As Double
    Dim local_score As Double
    
    Dim player_num As Byte
    Dim move As Byte
    
    ' Calculates Scores in the Board
    score = EvaluateBoard(board_table)
    
    If TTTLogic.IsAnyMoveValid(board_table) = False Or score = eMinMaxScores.COMPUTERWIN Or score = eMinMaxScores.HUMANWIN Then
        MinMaxAlghoritm = score
        Exit Function
    End If
    
    ' Place player number.
    If is_maximizer_turn = True Then
        player_num = GetCurrentPlayerNum()
        maximum_eval = eMinMaxScores.MAXIMUMCOMPUTER
        
    ElseIf is_maximizer_turn = False Then
        player_num = GetOppositePlayerNum()
        minimum_eval = eMinMaxScores.MAXIMUMHUMAN
        
    End If
    
    For move = LBound(board_table) To UBound(board_table)
        
        If board_table(move) = eBoardMove.Unoccupied Then
            board_table(move) = player_num
            local_score = MinMaxAlghoritm(board_table, depth + 1, Not is_maximizer_turn)

            If is_maximizer_turn = True Then
            
                board_table(move) = eBoardMove.Unoccupied
                maximum_eval = WorksheetFunction.Max(board_table, local_score, maximum_eval)

                MinMaxAlghoritm = maximum_eval

            ElseIf is_maximizer_turn = False Then

                board_table(move) = eBoardMove.Unoccupied
                minimum_eval = WorksheetFunction.Min(board_table, local_score, minimum_eval)

                MinMaxAlghoritm = minimum_eval

            End If

        End If
        
    Next move

End Function

Private Function EvaluateBoard(board_table As Variant) As Double
' Checks if some player has won in the simulation.
    
    If TTTLogic.HasPlayerWin(GetCurrentPlayerNum(), board_table) = True Then
        EvaluateBoard = eMinMaxScores.COMPUTERWIN
        
    ElseIf TTTLogic.HasPlayerWin(GetOppositePlayerNum(), board_table) = True Then
        EvaluateBoard = eMinMaxScores.HUMANWIN
                
    Else:
        EvaluateBoard = eMinMaxScores.TIE
    
    End If

End Function


