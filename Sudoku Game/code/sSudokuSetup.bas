Attribute VB_Name = "sSudokuSetup"
Option Explicit
' ---------------------------------------------------------------------------------------------------------------------------------------
' Module that stores all functions relaed to setting up the board from logical point
' WARNING! Does not makes sure that there is only one viable solution, however it accepts only one!
' ---------------------------------------------------------------------------------------------------------------------------------------

Public Function SetupSudokuBoard() As Integer()
    ' Main function of setting up the board to be played.
    ' It fills the empty board's first row with elements, then shuffles it, solves the game and removes some of elements.
    
    ' Accepts:
    '   None
    
    ' Returns:
    '   boardTable (2D array of integers that is full the board with removed some of the elements.

    Dim boardTable() As Integer
    Dim isSolvable As Boolean
    
    ReDim boardTable(SudokuConstans.FirstRow To SudokuConstans.LastRow, _
                                 SudokuConstans.FirstRow To SudokuConstans.LastRow)
    
    On Error GoTo err
    
    boardTable = ShuffleRandomRow(boardTable, 0) ' chosen the first row
    isSolvable = sSudokuSolver.CanSudokuBeSolved(boardTable)
    
    boardTable = sSudokuSolver.GetsolvedBoard()
    boardTable = RemoveRandomElementsFromBoard(boardTable, gGlobConfigs.GetNumToRemove())
    
    SetupSudokuBoard = boardTable
    Exit Function
    
err:
    err.Raise 1, "Sudoku Setup", "Cannot setup board correctly"

End Function

Private Function ShuffleRandomRow(ByRef boardTable() As Integer, numRow As Byte) As Integer()
    ' As name implies, shuffles randomly elements in numRow row of an 2D array.
    
    ' Accepts:
    '   boardTable [the 2D array of integers that will be modified]
    '   numRow [byte the number of row in which shuffling will have place]
    
    ' Returns:
    '   boardTable [the 2D array of integers]

    Dim r As Byte
    Dim randomIndex As Byte
    Dim temp As Byte
    
    For r = FirstRow To LastRow
        boardTable(numRow, r) = r + 1
    Next r

    ' Shuffles elements in the chosen row
    For r = FirstRow To LastRow
        randomIndex = Application.WorksheetFunction.RandBetween(FirstRow, LastRow)
        temp = boardTable(numRow, r)
        boardTable(numRow, r) = boardTable(numRow, randomIndex)
        boardTable(numRow, randomIndex) = temp
    Next r

    ShuffleRandomRow = boardTable

End Function

Private Function RemoveRandomElementsFromBoard(ByRef boardTable() As Integer, numToRemove As Integer) As Integer()
    ' As name implies remove random elements from 2D board with numToRemove quantity.
    
    ' Accepts:
    '   boardTable [2D table from which the elements will be removed]
    '   numToRemove [quantity of elements that will be removed from the board.
    
    ' Returns:
    '   boardTable [2D modified array]
    
    Dim numRow As Integer, numCol As Integer
    Dim indicesToRemove() As Variant
    Dim i As Integer
    
    numRow = LastRow
    numCol = LastRow
    
    ReDim indicesToRemove(1 To numToRemove)
    
    ' Fullfiles indices to remove with empty 2D arrays
    For i = 1 To numToRemove
        indicesToRemove(i) = Array(0, 0)
    Next i
    
    ' Fullfils indicies to remove with random elements.
    For i = 1 To numToRemove
        Dim rowRandomIndex As Integer
        Dim colRandomIndex As Integer
        Dim indexNum As Integer
            
        ' Adds element to indicesToRemove till it is unique
        Do While IsSubarrayInArray(indicesToRemove, indicesToRemove(i))
            rowRandomIndex = Application.WorksheetFunction.RandBetween(0, numRow)
            colRandomIndex = Application.WorksheetFunction.RandBetween(0, numCol)
            
            indicesToRemove(i) = Array(rowRandomIndex, colRandomIndex)
                
        Loop
        
    Next i
    
    ' Remove saved elements
    For i = 1 To numToRemove
        boardTable(indicesToRemove(i)(0), indicesToRemove(i)(1)) = 0
    Next i
    
    RemoveRandomElementsFromBoard = boardTable

End Function

Private Function IsSubarrayInArray(mainArray() As Variant, subArrayToSearch As Variant) As Boolean
    ' Helper function for RemoveRandomElementsFromBoard.
    ' Checks if position (subArrayToSearch) is already stored within mainArray (boardTable)
    
    ' Accepts:
    '   mainArray [array of arrays that stores values]
    '   subArrayToSearch [subarray to seach in mainArray]
    
    ' Returns:
    '   boolean [True means that subarray is already in array, false means otherwise]
    
    Dim i As Integer, count As Integer
    
    count = 0
    For i = LBound(mainArray) To UBound(mainArray)
        If mainArray(i)(0) = subArrayToSearch(0) And mainArray(i)(1) = subArrayToSearch(1) Then
            count = count + 1
            
            If count = 2 Then
                IsSubarrayInArray = True
                Exit Function
            End If
            
        End If
    Next i

    IsSubarrayInArray = False

End Function
