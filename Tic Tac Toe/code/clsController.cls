VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Board As FormsBoard
Private players As clsPlayerHandler

Private Property Set SetBoard(value As FormsBoard)
    
    Set Board = value
    
End Property

Private Property Get GetBoard()

    Set GetBoard = Board

End Property

Private Property Set SetPlayers(value As clsPlayerHandler)

    Set players = value

End Property

Private Property Get GetPlayers()

    Set GetPlayers = players
    
End Property

Public Function InitializeGame(settings As Scripting.Dictionary)
' Starts the Game.
    
    InitializeObjects
    
    players.InitializePlayers settings
    
    Board.InitializeBoard
    Board.PlacePlayerData players.GetPlayerObject(eTurn.CurrentTurn), players.GetPlayerObject(eTurn.NextTurn)
    
    MakeComputerMove eTurn.CurrentTurn
    
    Board.Show

End Function

Public Function RestartGame()
' Restarts the game by creating new Board Object

    Dim NewBoard As New FormsBoard
    
    Unload Board
    
    Set SetBoard = NewBoard
   
    Board.InitializeBoard
    Board.PlacePlayerData players.GetPlayerObject(eTurn.CurrentTurn), players.GetPlayerObject(eTurn.NextTurn)
    MakeComputerMove eTurn.CurrentTurn
    
    Board.Show
    
End Function

Public Function ExitGame()
' Finishes and exits the game
    
    Set players = Nothing
    
    Unload Board
    Set Board = Nothing
    
End Function

Public Function MakeMove(button_num As Byte)

    Dim current_turn As Byte
    
    current_turn = TTTLogic.CountNonEmptyBoardSpaces(Board.GetBoardTable())
    
    ' Place image and point in the table.
    Board.GetButtonsArray(button_num).PlaceImage (players.GetPlayerProperty(current_turn, Icon))
    Board.LetBoardTable button_num, players.GetPlayerProperty(current_turn, Number)

    ' Check whether there is any winner or tie.
    If TTTLogic.HasPlayerWin(players.GetPlayerProperty(current_turn, Number), Board.GetBoardTable()) = True Then
        TTTLogic.FinishGame players.GetPlayerObject(current_turn), True

    ElseIf TTTLogic.IsAnyMoveValid(Board.GetBoardTable()) = False Then
       TTTLogic.FinishGame players.GetPlayerObject(current_turn), False

    End If
    
    ' Prevents Automatic Error and makes computer move.
    If Not players Is Nothing Then
        MakeComputerMove current_turn + 1
    End If

End Function

Private Function MakeComputerMove(current_turn As Byte)
' Function responsible for Handling Computer Move.

    Dim ComputerMove As New clsComputerMoves
    Dim res As Byte

    If players.GetPlayerProperty(current_turn, PlayerType) = Enums.TranslatePlayerTypes(AI) Then
        
        ComputerMove.LetCurrentPlayerNum = players.GetPlayerProperty(current_turn, Number)
        ComputerMove.LetOppositePlayerNum = players.GetPlayerProperty(current_turn + 1, Number)
        
        res = ComputerMove.MakeMove(players.GetPlayerProperty(current_turn, ComputerType), Board.GetBoardTable())
        MakeMove res
        
    End If
    
    Set clsComputerMoves = Nothing
    
End Function

Private Function InitializeObjects()
    ' Factory Method that initializes all necessary objects in controller.
    
    Dim local_board As New FormsBoard
    Dim local_Players As New clsPlayerHandler
    
    Set SetBoard = local_board
    Set SetPlayers = local_Players
    
End Function


