Attribute VB_Name = "sSudokuSolver"
Option Explicit

' b in module name means "board"
' ---------------------------------------------------------------------------------------------------------------------------------------
' Module that handles solving the sudoku puzzle game.
' WARNING! Does not makes sure that there is only one viable solution, however it accepts only one!
' ---------------------------------------------------------------------------------------------------------------------------------------

' Module level variables
Private solvedBoard() As Integer ' Attribute that represents the state of the correctly solvedSudoku board.

Public Property Get GetsolvedBoard() As Integer()
    ' Getter for solvedBoard variable that stores the 2D arr of integers and represents correctly solvedBoard
    
    GetsolvedBoard = solvedBoard

End Property

Private Property Let LetsolvedBoard(value() As Integer)
    ' Letter for solvedBoard variable that stores the correctly solved Board.
    
    solvedBoard = value

End Property

' ---------------------------------------------------------------------------------------------------------------------------------------------------------
' ----------------------------------------------------------- ENTRY POINT ------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------------------------------------------

Public Function CanSudokuBeSolved(boardTable() As Integer) As Boolean
    ' A function that asses whether current sudoku can be possibly solved or not.
    
    ' Accepts:
    '    boardTable [value of a table that will be base for backtracking calculations]
    
    ' Returns:
    '    Bool [True if sudoku is solvable, False if it is not possible to solve it]
    
    ' Dim boardTable() As Integer
    Dim is_solvable As Boolean
    
    ' ReDim boardTable(GlobVar.FIRSTROW To GlobVar.LastRow, GlobVar.FIRSTROW To GlobVar.LastRow)
    
    If SolveSudoku(boardTable) Then
        CanSudokuBeSolved = True
    
    Else
        CanSudokuBeSolved = False
        
    End If

End Function

' ----------------------------------------------------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------- BACKTRACKING -- ---------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------------------------------------------------

Private Function SolveSudoku(ByRef boardTable() As Integer) As Boolean
    ' Recursive function using backtracking aghoritm to find a way to solve the Sudoku Puzzle.
    
    ' Accepts:
    '   boardTable()  [2D array of integers representing current board state of the game]
    
    ' Returns:
    '   bool [True puzzle can be solved, False - puzzle cannot be solved]
    
    Dim positions() As Variant
    Dim localRow As Byte, localColumn As Byte
    Dim digit As Byte
    
    positions = Array(0, 0) ' this variable stores the position of the next empty position

    ' Searches for next non empty position in the boardTable
    If Not FindNextEmptyLocation(boardTable, positions) Then
        ' Assigns finished boardTable to the module level variable.
        SolveSudoku = True
        Exit Function
    End If
    
    localRow = positions(0)
    localColumn = positions(1)
    
    ' Tests all possible digits in the positions
    For digit = SudokuConstans.firstvalidnum To SudokuConstans.lastValidNum
        If CanDigitBePlaced(boardTable, CByte(localRow), CByte(localColumn), digit) Then
            boardTable(localRow, localColumn) = digit
            
            If SolveSudoku(boardTable) Then
                solvedBoard = boardTable
                SolveSudoku = True
                Exit Function
            End If
            
            boardTable(localRow, localColumn) = 0
        
        End If
        
    Next digit
    
    SolveSudoku = False

End Function


Private Function FindNextEmptyLocation(ByRef boardTable() As Integer, positions As Variant) As Boolean
    ' Loops throught boardTable and searches for non empty location in in.
    
    ' Accepts:
    '   boardTable()  [2D array of integers representing current board state of the game]
    '   postitions [2 elements array representing last non taken position in the board]
    
    ' Returns:
    '   bool [True if there is to be found an empty place in the board, False if it is not]

    Dim r As Byte, c As Byte
    
    For r = SudokuConstans.FirstRow To SudokuConstans.LastRow
        For c = SudokuConstans.FirstRow To SudokuConstans.LastRow
        
            If boardTable(r, c) = 0 Then
                positions(0) = r
                positions(1) = c
                FindNextEmptyLocation = True
                Exit Function
            End If
            
        Next c
    Next r
    
    FindNextEmptyLocation = False

End Function

Private Function CanDigitBePlaced(ByRef boardTable() As Integer, row As Byte, column As Byte, digit As Byte) As Boolean
    ' Summarive function that checks whether in boardTable state of the board it is possible to place digit in the position boardTable(row, column) - according to Sudoku Rules.
    
    ' Accepts:
    '   boardTable() [2D array representing current state of the board]
    '   row  [byte representing current row]
    '   column  [column number that represents the current column number in the board]
    '   digit [number that should be placed in the position]
    
    ' Returns:
    '   bool [True it is possible to place in boardTable(row, column) a digit, otherwise false.
    
    
    If Not IsDigitUsedInRow(boardTable, row, digit) And _
    Not IsDigitUsedInColumn(boardTable, column, digit) And _
    Not IsDigitUsedInSquare(boardTable, row, column, digit) Then
        CanDigitBePlaced = True
        Exit Function
    End If
    
    CanDigitBePlaced = False
    
End Function

Private Function IsDigitUsedInRow(ByRef boardTable() As Integer, row As Byte, digit As Byte) As Boolean
    ' Checks whether a digit is placed currently in boardTable(row, ...) row.
    
    ' Accepts:
    '   boardTable() [2D array representing current state of the board]
    '   row [byte representing row that will be tested]
    '   digit [byte representing the number that function will check for]
    
    ' Returns:
    '   bool [True if digit is placed in boardTable(row, ...) otherwise False
    
    Dim r As Byte
    
    For r = SudokuConstans.FirstRow To SudokuConstans.LastRow
        If boardTable(row, r) = digit Then
            IsDigitUsedInRow = True
            Exit Function
        End If
    Next r
    
    IsDigitUsedInRow = False
        
End Function

Private Function IsDigitUsedInColumn(ByRef boardTable() As Integer, column As Byte, digit As Byte) As Boolean
    ' Checks whether a digit is placed currently in boardTable(..., column).
    
    ' Accepts:
    '   boardTable() [2D array representing current state of the board]
    '   column[byte representing column that will be tested]
    '   digit [byte representing the number that function will check for]
    
    ' Returns:
    '   bool [True if digit is placed in boardTable(..., column) otherwise False
    
    Dim c As Byte
    
    For c = SudokuConstans.FirstRow To SudokuConstans.LastRow
        If boardTable(c, column) = digit Then
            IsDigitUsedInColumn = True
            Exit Function
        End If
    Next c
    
    IsDigitUsedInColumn = False

End Function

Private Function IsDigitUsedInSquare(ByRef boardTable() As Integer, row As Byte, column As Byte, digit As Byte) As Boolean
    ' Checks whether a number digit is stored in the current 3x3 square.

     ' Checks whether a digit is placed currently in boardTable(..., column).
    
    ' Accepts:
    '   boardTable() [2D array representing current state of the board]
    '   row [byte representing row from which nearest 3x3 square will searched for]
    '   column [byte representing column from which nearest 3x3 square will searched for]
    '   digit [byte representing the number that will be searched in the current 3x3 square.]
    
    ' Returns:
    '   bool [True if digit is placed in boardTable(..., column) otherwise False
    
    Dim xSquare As Byte, ySquare As Byte
    Dim r As Byte, c As Byte

    '  Indicates current square position
    xSquare = row - (row Mod SudokuConstans.squareRowNum)
    ySquare = column - (column Mod SudokuConstans.squareRowNum)
    
    For r = SudokuConstans.FirstRow To SudokuConstans.squareLastRow
        For c = SudokuConstans.FirstRow To SudokuConstans.squareLastRow
            
            If boardTable(r + xSquare, c + ySquare) = digit Then
                IsDigitUsedInSquare = True
                Exit Function
            End If
            
        Next c
    Next r
    
    IsDigitUsedInSquare = False

End Function

