Attribute VB_Name = "Enums"
' This module contains all constans used in the Tic Tac Toe Project.
Option Explicit
Public Enum eSettings
' TRANSLATION FUNCTION -> TranslateSettings

    Player1Name = 1
    Player1Icon = 2
    Player1Number = 3
    Player1Type = 4

    Player2Name = 5
    Player2Icon = 6
    Player2Number = 7
    Player2Type = 8
    
    player1ComputerType = 9
    player2ComputerType = 10
    
End Enum

Public Enum eBoardMove
    
    Player1Move = 1
    Player2Move = 2
    Unoccupied = 0
    MiddlePoint = 4
    
End Enum

Public Enum eTurn
    
    CurrentTurn = 0
    NextTurn = 1

End Enum

Public Enum ePlayerProperty
    
    Name = 1
    Icon = 2
    Number = 3
    PlayerWon = 4
    PlayerType = 5
    ComputerType = 6

End Enum

Public Enum ePlayerTypes
' Translation Function -> TranslatePlayerTypes()

    AI = 1
    HUMAN = 2
    
End Enum

' -----------------------------------------------
' COMPUTER PLAYER CONSTANS
' -----------------------------------------------

Public Enum eMinMaxScores

    COMPUTERWIN = 10
    HUMANWIN = -10
    TIE = 0
    MAXIMUMCOMPUTER = -1000
    MAXIMUMHUMAN = 1000

End Enum

Public Enum eComputerMode
    
    random = 1
    UNBEATABLE = 2
    mixed = 3
    
End Enum

' -------------------------------------------------------------------------------
' TRANSLATIONS OF ENUMS TO ANOTHER DATA TYPES
' -------------------------------------------------------------------------------

Public Function TranslatePlayerTypes(ePlayerType As ePlayerTypes) As String
    
    Dim TranslationArray() As Variant
    Dim num As Byte
    
    TranslationArray = Array("", "AI", "HUMAN")
    
    For num = LBound(TranslationArray) To UBound(TranslationArray)
        If num = ePlayerType Then
            TranslatePlayerTypes = TranslationArray(num)
            Exit Function
        End If
    Next num

End Function

Public Function TranslateSettings(SettingNum As eSettings) As String
    
    Dim TranslationArray() As Variant
    Dim num As Byte
    
    TranslationArray = Array("", "player1name", "player1icon", "player1number", "player1type", "player2name", _
                                               "player2icon", "player2number", "player2type", "player1ComputerType", "player2ComputerType")
    
    For num = LBound(TranslationArray) To UBound(TranslationArray)
        If num = SettingNum Then
            TranslateSettings = TranslationArray(num)
            Exit Function
        End If
    Next num

End Function

Public Function TranslateComputerMode(eComputer As eComputerMode) As String
    
    Dim TranslationArray() As Variant
    Dim num As Byte
    
    TranslationArray = Array("", "RANDOM", "UNBEATABLE", "MIXED")
    
    For num = LBound(TranslationArray) To UBound(TranslationArray)
        If num = eComputer Then
            TranslateComputerMode = TranslationArray(num)
            Exit Function
        End If
    Next num
    
End Function
