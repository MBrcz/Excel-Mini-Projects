VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DataClass that holds only properties of the Player object.
Option Explicit
Private Icon As String ' Icon of player
Private Name As String ' Name of player
Private Number As Byte ' The representation in the board.
Private WonGames As Byte ' How many players won durning this session
Private PlayerType As String ' AI or Human
Private ComputerType As String ' Difficulity of a AI player (if it's AI)

Public Property Let LetComputerType(value As String)
    
    If GetPlayerType() = Enums.TranslatePlayerTypes(AI) Then
        ComputerType = value
    
    Else
        ComputerType = ""
    End If

End Property

Public Property Get GetComputerType()
    
    GetComputerType = ComputerType

End Property

Public Property Let LetPlayerType(value As String)

    PlayerType = value

End Property

Public Property Get GetPlayerType() As String

    GetPlayerType = PlayerType
    
End Property

Public Property Let LetNumber(value As Byte)

    Number = value
    
End Property

Public Property Get GetNumber() As Byte

    GetNumber = Number
    
End Property

Public Property Let LetName(value As String)

    Name = value
    
End Property

Public Property Get GetName() As String

    GetName = Name
    
End Property

Public Property Let LetIcon(icon_name As String)

    Icon = icon_name
    
End Property

Public Property Get GetIcon() As String

    GetIcon = Icon
    
End Property

Public Property Let LetWonGames(value As Byte)

    WonGames = value
    
End Property

Public Property Get GetWonGames() As Byte

    GetWonGames = WonGames
    
End Property
