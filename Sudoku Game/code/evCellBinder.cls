VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "evCellBinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --------------------------------------------------------------------------------------------------------------------
' Purpouse of this class is to store and bind all cells in the GUIBoard Forms. ---------------------------------------
' By the design, this class should only communicate with mController module. -----------------------------------------
' --------------------------------------------------------------------------------------------------------------------

Public WithEvents cell As MSForms.textbox ' textbox to storage
Attribute cell.VB_VarHelpID = -1
Public cellChange As Boolean ' prevents triggering cell_change event durning initialization phase
Public cellRow As Byte ' row of a cell in a grid
Public cellColumn As Byte ' column of a cell in a grid

' -------------------------------------------------------------
' ------------------ CALLBACK --------------------------
' -------------------------------------------------------------
Private Sub cell_change()
    ' Callback function for handling the change event of the cell object.

    ' Accepts:
    '   None
    ' Returns:
    '   None

    If cellChange Then
        If cell.value = "" Then
            HandleEmptyCell
        ElseIf Not IsInputValid() Then
            ClearCellValue
        Else
            CheckInputAndHandleResult
        End If
        
        mController.HasPlayerLost
    End If
    
End Sub

' -----------------------------------------------------------------
' ----------------------- FUNCTIONS -----------------------
' -----------------------------------------------------------------

Private Sub HandleEmptyCell()
    ' Function that changes the color to the white and unlocks whole board.
    ' Usefull when player set the wrong number and clears the cell.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None

    cell.BackColor = eEnums.GetRGBColor(SudokuColors.white)
    mController.UnblockWholeBoard
    
End Sub

Private Sub ClearCellValue()
    ' As name implies, clears the value of the to the ""

    cell.value = ""
    
End Sub

Private Sub CheckInputAndHandleResult()
    ' Checks input of the cell and handles it result [as name implies]
    
    ' Accepts:
    '   None
    ' Returns:
    '   None

    If Not mController.CheckInput(cellRow, cellColumn) Then
        HandleIncorrectInput
    Else
        HandleCorrectInput
    End If
    
End Sub

Private Sub HandleIncorrectInput()
    ' Changes color of a cell to red, blocks whole board and update lives.
    ' Used where user thrown valid but not correct input.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None

    cell.BackColor = eEnums.GetRGBColor(SudokuColors.red)
    mController.BlockWholeBoard
    mController.UpdateLives True
    cell.Enabled = True
    cell.SetFocus
    
End Sub

Private Sub HandleCorrectInput()
    ' Updates the counter, disables the cell to the edition and checks whether a player could have won a game.
    ' Used when user throws a valid and correct input.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None

    mController.UpdateCounter
    cell.Enabled = False
    mController.HasPlayerWon
    
End Sub

Private Function IsInputValid() As Boolean
    ' Checks if input set by user is possible to be valid via Sudoku rules.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    Dim i As Byte
    
    For i = SudokuConstans.firstvalidnum To SudokuConstans.lastValidNum
        If cell.value = i Then
            IsInputValid = True
            Exit Function
        End If
    Next i
        
    IsInputValid = False

End Function

