Attribute VB_Name = "mController"
Option Explicit

' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' This is controller module. The basic idea is that it is a central module of an application that holds all neccessary methods for bindings of cells would be possible.
' It communicates with all other modules in the application.

' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub StartGame()
    ' Starts completely new game, usable only once.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    Application.Visible = False
    Dim gui As New UserForm1
    
    gui.Show
    Application.Visible = True
    
End Sub

Public Function loadController(boardObject() As evCellBinder, ufBoard As UserForm1)
    ' Init method of the module. At the start of the application using factory method it sets all variables for programme to work.
    
    ' Note only used by: UserForm1.RestartGame method.
    
    ' Accepts:
    '   boardObject [2D array of evCellBinder objects, used to store all cells in the board (as TextBox objects). See more: evCellBinder class]
    '   ufBoard [UserForm Board as an object - necessary for restarting game and controlling the supportFrame]
    
    ' Returns:
    '   None
    
    moduleFactoryController boardObject, ufBoard
    
End Function

Public Function UpdateLives(hasScoreBeenFailed As Boolean)
    ' Helper method for updating the lives in the game. See more: bSupport.UpdateCurrentLives.
    
    ' Accepts:
    '   hasScoreBeenFailed [bool, whether input provided by user is correct]

    ' Returns:
    '   None
    
    bSupport.UpdateCurrentLives hasScoreBeenFailed
    
End Function

Public Function UpdateCounter()
    ' Helper method that updates the counter of a filled by user cells. See more: bSupport.UpdateTheCounter
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    bSupport.UpdateTheCounter

End Function

Public Function BlockWholeBoard()
    ' Blocks the possibility of changing the values of a cells by user. See more: bBoard.DisableBoardEdition
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    bBoard.DisableBoardEdition

End Function

Public Function UnblockWholeBoard()
    ' Enables the possibility of the edition of a non empty cells in the board see more: bBoard.EnableBoard
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    bBoard.EnableBoard
    
End Function

Public Function CheckInput(xPos As Byte, yPos As Byte) As Boolean
    ' Checks whether cell input is correct with the completedTable. See more: bBoard.IsBoard.Correct
    
    ' Accepts:
    '   xPos [byte, x posiition in the grid of a cell]
    '   yPos [byte, y position in the grid of a cell]
    ' Returns:
    '   Bool [True - position is correct, False - position is incorrect]
    
    CheckInput = bBoard.IsBoardCorrect(xPos, yPos)

End Function

Public Function HasPlayerWon()
    ' Checks whether a player could won a game in current state See more: bSupoprt.DidPlayerWon
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
        
    bSupport.DidPlayerWon

End Function

Public Function HasPlayerLost()
    ' Checks whether a player could lose a game in current state See more: bSupoprt.DidPlayerLost
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    bSupport.DidPlayerLost
    
End Function

Private Function moduleFactoryController(boardObject() As evCellBinder, ufBoard As UserForm1)
    ' Factory method for current module. It sets up all neccessary properties to another modules in order for programme to work.
    
    ' Accepts:
    '   boardObject [2D array of evCellBinder objects, used to store all cells in the board (as TextBox objects). See more: evCellBinder class]
    '   ufBoard [UserForm Board as an object - necessary for restarting game and controlling the supportFrame]
    
    ' Returns:
    '   None
    
    bBoard.LetguiBoard boardObject
    bBoard.LetcompletedBoard = sSudokuSetup.SetupSudokuBoard()
    
    bBoard.PopulateBoardWithNumbers
    bBoard.LetcompletedBoard = sSudokuSolver.GetsolvedBoard
    
    BlockWholeBoard
    UnblockWholeBoard
    
    bSupport.LetliveCounter = SudokuGameplay.defaultLivesNum ' see more eEnums
    bSupport.LetstartTimer = Timer
    Set bSupport.SetufBoard = ufBoard
    
    bSupport.StartCountingTime
    
    UpdateLives False
    UpdateCounter
    
End Function


