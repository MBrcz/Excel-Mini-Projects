Attribute VB_Name = "eEnums"
Option Explicit

' e in module name means "Enum"
' --------------------------------------------------------------------------------------------------------------------------------------------
' This module contains enums (constans related to the project)
' I would really love to switch names of all constans to CAMEL CASE but for some reason it is not possible.
' --------------------------------------------------------------------------------------------------------------------------------------------

Public Enum SudokuConstans
    ' This enum function represents all constans related to the sudoku game
    FirstRow = 0 ' num of first square in 9x9 board (starts from 0 due to the fact that arrays in VBA starts at 0)
    LastRow = 8 'last num of square in 9x9 board
    
    firstvalidnum = 1 ' num of the first valid input possible to place in the board
    lastValidNum = 9 ' num of the last valid input possible to place in the board
    sumOfSquares = 81 ' the total value of cells in 9x9 board
    
    squareRowNum = 3
    squareLastRow = 2
    
    ' defaultLivesNum = 3 ' default quantity of lives
    
End Enum
    
Public Enum SudokuGameplay
    ' Constans related to the sudoku gameplay enum.
    
    defaultLivesNum = 3
    
End Enum

Public Enum SudokuColors
    ' Helper enum towards GetRGB colors fucntion.
    
    red = 1
    white = 2
    
End Enum

Function GetRGBColor(ByVal color As SudokuColors) As Long
    ' Gets colour as integer and returns it as a RGB num.
    
    ' Accepts:
    '   color [enum - the number that represents the color]
    
    ' Returns:
    '   long [represents the rgb number as the code]
    
    Select Case color
        Case SudokuColors.red
            GetRGBColor = RGB(255, 0, 0)
        Case SudokuColors.white
            GetRGBColor = RGB(255, 255, 255)
        Case Else
            GetRGBColor = 0
    End Select
End Function
    
