Attribute VB_Name = "gGlobConfigs"
Option Explicit
' ------------------------------------------------------------------------------
' This module contains configs that user can change in the game via sheet.
' Alas, there is only one here
' -------------------------------------------------------------------------------

Const CONFIGSHEET = "_SUDOKU_GAME_"
Public Function GetNumToRemove() As Integer
    ' Gets from the worksheet quantity of elements which will be removed from completed board.
    
    ' Accepts:
    '   None
    ' Returns:
    '   None
    
    Dim v As Integer
    
    On Error GoTo err
    
    v = CInt(Sheets(CONFIGSHEET).Cells(3, 2).value)
    If v >= 1 And v <= 81 Then
        GetNumToRemove = v
        Exit Function
    End If
    
err:
    MsgBox "Cannot load correctly quantity of numbers to remove, therefore default value 40 is set.", vbCritical
    GetNumToRemove = 40

End Function

