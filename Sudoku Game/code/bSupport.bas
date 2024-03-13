Attribute VB_Name = "bSupport"
Option Explicit

' -----------------------------------------------------------------------------------------
' This module contains all data related to the logical functions that are bound to the board and programme
' -----------------------------------------------------------------------------------------

Private ufBoard As UserForm1 ' Module level property that represents whole UserForm object
Private liveCounter As Integer ' Module level property that represents integer as liveCounter
Private startTimer As Double ' Module level property that is set to count the time in which the user solves the Sudoku.

' --------------------------------------------------------------------------------------------------------------------------
' --------------------------------------------- PROPERTIES --------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------------------
' liveCounter
Public Property Let LetliveCounter(value As Integer)
    
    liveCounter = value

End Property

Private Property Get GetliveCounter() As Integer
    
    getlivecoutner = liveCounter

End Property

' ufBoard
Public Property Set SetufBoard(value As UserForm1)
    
    Set ufBoard = value

End Property

Private Property Get GetufBoard()
        
    Set GetufBoard = ufBoard

End Property

' startTimer
Public Property Let LetstartTimer(value As Double)
    
    startTimer = value

End Property

Private Property Get GetstartTimer()
    
    GetstartTimer = startTimer

End Property

' ---------------------------------------------------------------------------------------------------------------------------------------
' ----------------------------------------------- PUBLIC FUNTIONS ------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------------------------
Public Function DidPlayerLost()
    ' Tests whether player could have lost a game.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    If liveCounter = 0 Then
        FinishLostGame
    End If

End Function

Public Function DidPlayerWon()
    ' Tests whether player could have win a game.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    If bBoard.CountFilledCells = SudokuConstans.sumOfSquares Then
        FinishWonGame
    End If

End Function

Public Function StartCountingTime()
    ' Start counting time in the game.
    
    ' Accpets:
    '   None
    ' Returns:
    '   None
    
    LetstartTimer = Timer

End Function

Public Function FinishCountingTime() As Double
    ' Finishes counting the time in the Game
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    Dim finishTimer As Double
    finishTimer = Timer
    
    FinishCountingTime = Round(finishTimer - startTimer, 2)

End Function

Public Function UpdateCurrentLives(isScoreFailed As Boolean)
    ' Updates current live count in the GUI
    
    ' Accepts:
    '   isScoreFailes [bool - True there is neccesity to decrease the liveCount, False there is no]
    
    ' Returns:
    '   None
    
    If isScoreFailed Then
        liveCounter = liveCounter - 1
    End If
    UpdateLivesView liveCounter
    
End Function

Public Function UpdateTheCounter()
    ' Updates the counter quantity in the board object
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    Dim count As Byte
    Dim SupportFrame As MSForms.frame
    
    Set SupportFrame = ufBoard.controls("supportFrame")
    count = bBoard.CountFilledCells()
    
    SupportFrame.controls("Counter").text = CStr(SudokuConstans.sumOfSquares - count)

End Function

' ----------------------------------------------------------------------------------
' PRIVATE FUNCTIONS ---------------------------------------------------
' ----------------------------------------------------------------------------------

Private Function UpdateLivesView(livesNum As Integer)
    ' Updates the liveCounter in the GUI
    
    ' Accepts:
    '   livesNum [quantity of squares which color would be set to red]
    ' Returns:
    '   None

    Dim livesBoxNames() As Variant
    Dim SupportFrame As MSForms.frame
    Dim nameNum As Byte
    
    Set SupportFrame = ufBoard.controls("supportFrame")
    livesBoxNames = Array("LiveBox1", "LiveBox2", "LiveBox3")

    For nameNum = LBound(livesBoxNames) To UBound(livesBoxNames)
        SupportFrame.controls(livesBoxNames(nameNum)).BackColor = eEnums.GetRGBColor(SudokuColors.white)
        
        If nameNum < livesNum Then
            SupportFrame.controls(livesBoxNames(nameNum)).BackColor = eEnums.GetRGBColor(SudokuColors.red)
        End If

    Next nameNum
    
End Function

Private Function FinishLostGame()
    ' Finishes the game and restarts it whenever user has lost;
    
    ' Accept:
    '   None
    ' Returns:
    '   None
    
    Dim msg As String
        
    msg = "You have Lost! Try again! " & GetTimeMessage()
    
    MsgBox msg
    RestartGame

End Function

Private Function FinishWonGame()
    ' Finishes the game and restarts it whenever user has won, which probably will never happen ;x
    
    ' Accept:
    '   None
    ' Returns:
    '   None
    
     Dim msg As String
     msg = "You have Won! Congratulations! " & GetTimeMessage()
    
     MsgBox msg
     RestartGame

End Function

Private Function GetTimeMessage() As String
    
    Dim time As Double
    time = bSupport.FinishCountingTime()
    
    GetTimeMessage = " You have fought for " & CStr(time) & " seconds with this Sudoku!"

End Function

Private Function RestartGame()
    ' Restarts the whole Game again.
    
    ' Accept:
    '   None
    ' Returns:
    '   None
    
    liveCounter = SudokuGameplay.defaultLivesNum
    ufBoard.RestartBoard

End Function

