Attribute VB_Name = "bBoard"
Option Explicit
' This module contains all related to logic of the Board functions.

' ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Theoretically this architecture is wrong and one should not kept variables bound to the modules as globals, but in this case this solution is sufficient.
' Also for some some reason durning refreshing SolvingBoard via Backtracking function the excel freezed causing a crash.
' Do not know whether is it a Excel's fault or my code so i chose this [anti-pattern] solution
' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private guiBoard() As evCellBinder ' Current state of board in the gui
Private completedBoard() As Integer ' The generated and solved board, kept to be compared with guiBoard

Public Function LetguiBoard(value() As evCellBinder)
    ' This function is a masquared Let property function. Is is sole purpose is to set the guiBoard property variable at a module level.
    ' By some VBA magic, traditional Setter&Let showed to be innefective here.
    
    Dim c As Byte, r As Byte
  
    ReDim guiBoard(LBound(value, 1) To UBound(value, 1), LBound(value, 1) To UBound(value, 1))
    
    For r = LBound(value, 1) To UBound(value, 1)
        For c = LBound(value, 1) To UBound(value, 1)
            Set guiBoard(r, c) = value(r, c)
        Next c
    Next r

End Function

Public Property Get GetguiBoard()
    ' Getter for the representation of GUIBoard object.
    GetguiBoard = guiBoard
End Property

Public Property Let LetcompletedBoard(value() As Integer)
    ' Setter for the arr() [integer, integer] that represents solved Sudoku game.
    completedBoard = value
End Property

Private Property Get GetcompletedBoard()
    ' Getter for completedBoard variable.
    GetcompletedBoard = completedBoard
End Property

' ------------------------------------------------------------------------------------------------------------------------------------------
' ------------------------------------------------ BOARD FUNCTIONS -----------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------------------

Public Function IsBoardCorrect(xPos As Byte, yPos As Byte) As Boolean
    ' Checks whether the cell set at completedBoard and in the guiBoard are the same.
    
    ' Accepts:
    '   xPos [byte] - row array position to be checked.
    '   yPos [byte] - column array position to be checked
    
    ' Returns:
    '   bool [True means that positions of both boards are equal, false otherwise]
    
    Dim areEqual As Boolean
    
    areEqual = CStr(completedBoard(xPos, yPos)) = CStr(guiBoard(xPos, yPos).cell.value)
    If areEqual Then
        IsBoardCorrect = True
    
    ElseIf Not areEqual Then
        IsBoardCorrect = False
    
    End If

End Function

Public Function CountFilledCells() As Integer
    ' Counts all cells in the board that are not empty and their background is not RED.
    ' Used as win condition and pure information for player of his/hers progress of the game.
    
    'Accepts:
    '   None
    
    ' Returns:
    '   Integer (cells that are filled and not incorrect)
    
    Dim r As Byte, c As Byte
    Dim count As Integer
    
    On Error GoTo err
    count = SudokuConstans.FirstRow
    For r = SudokuConstans.FirstRow To SudokuConstans.LastRow
        For c = SudokuConstans.FirstRow To SudokuConstans.LastRow
            If guiBoard(r, c).cell.value <> "" And guiBoard(r, c).cell.BackColor <> GetRGBColor(SudokuColors.red) Then
                count = count + 1
            End If
        Next c
    Next r
    
    CountFilledCells = count
    
    Exit Function

err:
    err.Raise 3, "BoardFunctions", "Counting cells has been failed failed."
    

End Function

Public Function DisableBoardEdition()
    ' Makes impossible to edit all cells in the guiBoard object.
    ' Used mostly in evButtonClick event for preventing a player from modifing the correct cell.
    
    'Accepts:
    '   None
    'Returns:
    '   None
    
    Dim r As Byte, c As Byte
    
    On Error GoTo err
    For r = SudokuConstans.FirstRow To SudokuConstans.LastRow
        For c = SudokuConstans.FirstRow To SudokuConstans.LastRow
            guiBoard(r, c).cell.Enabled = False
        Next c
    Next r
    
    Exit Function

err:
    err.Raise 3, "BoardFunctions", "Switching states of the cells was failed."

End Function

Public Function EnableBoard()
    ' Makes possible to change caption of cell textboxes, also removes the initialization mode via cellchange attribute.
    
    'Accepts:
    '   None
    'Returns:
    '   None
    
    Dim r As Byte, c As Byte
    
    On Error GoTo err
    
    For r = SudokuConstans.FirstRow To SudokuConstans.LastRow
        For c = SudokuConstans.FirstRow To SudokuConstans.LastRow
            
            If guiBoard(r, c).cell.value = "" Then
                guiBoard(r, c).cell.Enabled = True
                guiBoard(r, c).cellChange = True
            End If
            
        Next c
    Next r
    
    Exit Function

err:
    err.Raise 3, "BoardFunctions", "Enabling cells edition was failed."

End Function

Public Function PopulateBoardWithNumbers()
    ' Places the generated numbers in the guiBoard object. Used in initialization of programme only.
    
    'Accepts:
    '   None
    'Returns:
    '   None
    
    Dim r As Byte, c As Byte
    
    On Error GoTo err
    For r = SudokuConstans.FirstRow To SudokuConstans.LastRow
        For c = SudokuConstans.FirstRow To SudokuConstans.LastRow
            
            If completedBoard(r, c) <> "0" Then
                guiBoard(r, c).cell.value = completedBoard(r, c)
            ElseIf completedBoard(r, c) = "0" Then
                guiBoard(r, c).cell.value = ""
            End If
            
        Next c
    Next r
    
    Exit Function
err:
    err.Raise 3, "BoardFunctions", "Populating board with numbers was failed"

End Function

Public Function printBoard(ByRef boardTable() As Integer)
    ' Function for debugging purpouses only , prints a board as an string of bytes

    ' Accepts:
    '   boardTable() [2D array representing board]
    
    ' Returns:
    '   None

    Dim r As Byte, c As Byte
    Dim msg As String

    For r = GlobVar.FirstRow To GlobVar.LastRow
        msg = ""
        For c = GlobVar.FirstRow To GlobVar.LastRow
            msg = msg & boardTable(r, c) & ", "
        Next c
        Debug.Print msg
    Next r

End Function
