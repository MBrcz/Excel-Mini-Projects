Attribute VB_Name = "utils"
Option Explicit
' Module that contains all utility methods

Public Function ExtractNumbersFromString(text As String) As Integer
    ' As name implies extracts the digits from a string.
    
    ' Accepts:
    '   text [string from which extraction will happen]
    
    ' Return
    '   numLetter [integer the number that is extracted].
  
    Dim letter As String
    Dim numLetter As Byte
    Dim numInString As String
    
    On Error GoTo err
    
    ' Find number in string
    numInString = ""
    For numLetter = 1 To Len(text)
        letter = Mid(text, numLetter, 1)
        
        If IsNumeric(letter) Then
            numInString = numInString & letter
        End If
    
    Next numLetter
    
    ExtractNumbersFromString = CInt(numInString)
    Exit Function
    
err:
    err.Raise 0, "utils", "Position or number probably does not exist"
    
End Function

Public Function SplitTheNumberIntoPositions(ByVal number As Integer) As Variant
    ' Splits the numbers (integers) into positions of 9x9 array. For instance num 10 is converted to (1, 1).
    ' Used for initialization of the board.
    
    ' Accepts:
    '   number - number that will be split
    
    ' Returns:
    '   position (variant with numbers)
    
    Dim position() As Variant
    ReDim position(0 To 1)
    
    position(0) = number \ SudokuConstans.lastValidNum
    position(1) = number Mod SudokuConstans.lastValidNum
    
    SplitTheNumberIntoPositions = position

End Function
