VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlayerHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private player1 As clsPlayer
Private player2 As clsPlayer

Private Property Set SetPlayer1(value As clsPlayer)
    
    Set player1 = value

End Property

Private Property Set SetPlayer2(value As clsPlayer)

    Set player2 = value
    
End Property

Private Property Get GetPlayer1()

    Set GetPlayer1 = player1
    
End Property

Private Property Get GetPlayer2()

    Set GetPlayer2 = player2
    
End Property

Public Function InitializePlayers(settings As Scripting.Dictionary)
' Initializes players and assigns to them attributes.
    
    Dim player1 As New clsPlayer
    Dim player2 As New clsPlayer
    
    With player1
        .LetNumber = settings(Enums.TranslateSettings(Player1Number))
        .LetName = settings(Enums.TranslateSettings(Player1Name))
        .LetIcon = settings(Enums.TranslateSettings(Player1Icon))
        .LetWonGames = eBoardMove.Unoccupied
        .LetPlayerType = settings(Enums.TranslateSettings(Player1Type))
        .LetComputerType = settings(Enums.TranslateSettings(player1ComputerType))
    End With
    
    With player2
        .LetNumber = settings(Enums.TranslateSettings(Player2Number))
        .LetName = settings(Enums.TranslateSettings(Player2Name))
        .LetIcon = settings(Enums.TranslateSettings(Player2Icon))
        .LetWonGames = eBoardMove.Unoccupied
        .LetPlayerType = settings(Enums.TranslateSettings(Player2Type))
        .LetComputerType = settings(Enums.TranslateSettings(player2ComputerType))
    End With
    
    Set SetPlayer1 = player1
    Set SetPlayer2 = player2
    
End Function

Public Function GetPlayerProperty(game_turn As Byte, propertyName As ePlayerProperty) As Variant
    ' Getter for getting current or next player property (depends on propertyName argument)
    ' Game turn mod 2 = 0 means current player; mod 2 = 1 means next player
    
    Dim player As clsPlayer
    Dim property_value As Variant
    
    Set player = GetPlayerObject(game_turn)
    
    Select Case propertyName

        Case Icon
            property_value = player.GetIcon()

        Case Number
            property_value = player.GetNumber()

        Case PlayerWon
            property_value = player.GetWonGames()

        Case Name
            property_value = player.GetName()

        Case PlayerType
            property_value = player.GetPlayerType()
            
        Case ComputerType
            property_value = player.GetComputerType()

    End Select

    GetPlayerProperty = property_value

End Function

Public Function GetPlayerObject(game_turn As Byte) As clsPlayer
' Getter for player object.
    
    If game_turn Mod 2 = eTurn.CurrentTurn Then
        Set GetPlayerObject = GetPlayer1()
        
    ElseIf game_turn Mod 2 = eTurn.NextTurn Then
        Set GetPlayerObject = GetPlayer2()
        
    End If

End Function




