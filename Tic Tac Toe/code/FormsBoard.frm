VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormsBoard 
   Caption         =   "UserForm1"
   ClientHeight    =   6720
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6120
   OleObjectBlob   =   "FormsBoard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormsBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BoardTable() As Variant
Private ButtonsArray() As clsButtons

Public Property Get GetBoardTable()
    
    GetBoardTable = BoardTable()

End Property

Public Function LetBoardTable(position As Byte, value As Byte)
' This should be classified and treated like a property Let statement, but vba does not allow to use few arguments in setters.

    BoardTable(position) = value
    
End Function

Public Property Get GetButtonsArray()
    
    GetButtonsArray = ButtonsArray

End Property

Private Function SetButtonsArray(position As Byte, value As clsButtons, Button As MSForms.CommandButton)
' This should be classified and treated like a property Let statement, but vba does not allow to use few arguments in setters.

    ReDim Preserve ButtonsArray(position)
    Set ButtonsArray(position) = value
    Set ButtonsArray(position).ButtonObject = Button
    
End Function

Public Function InitializeBoard()
    
    Me.Caption = "Tic Tac Toe"
    
    ' Places TTT Window in the middle of the screen.
    Me.Top = utils.GetMiddlePointInScreen()(0)
    Me.Left = utils.GetMiddlePointInScreen()(1)
    
    InitializeBinds
    InitializeBoardTable
    
End Function

Public Function PlacePlayerData(player1 As clsPlayer, player2 As clsPlayer)

    Me.Player1Name = player1.GetName()
    Me.Player1Won = player1.GetWonGames()
    
    Me.Player2Name = player2.GetName()
    Me.Player2Won = player2.GetWonGames()

End Function

Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
' Event Handler for quitting the UserForm.

   If CloseMode = vbFormControlMenu Then
        Main.ExitGame
    End If
End Sub

Private Function InitializeBoardTable()
    
    Dim num As Byte
    
    ButtonsArray = GetButtonsArray()
    
    For num = LBound(ButtonsArray) To UBound(ButtonsArray)
        ReDim Preserve BoardTable(num)
        LetBoardTable num, eBoardMove.Unoccupied
    Next num

End Function

Private Function InitializeBinds()
' Binds callback to each button in Forms.

    Dim btn As MSForms.CommandButton
    Dim num As Byte
    
    num = 0
    Do While num < 9
        Set btn = FindButtonByNumber(CByte(num))
        SetButtonsArray CByte(Right(btn.Name, 1)), New clsButtons, btn
        num = num + 1
    Loop
  
End Function

Private Function FindButtonByNumber(button_num As Byte) As Control
' Each button that needs to be bound in ButtonsArray(), every button contains it's unique name, that has unique number at the end.
' For some unknown reasons, using convencional methods like (for each button: button name matches bind ) (PSEUDOCODE)
' Does not allow to bind buttons correctly, so i am forced to use this monstrocity instead.
    
    Dim Button As Control
    
    For Each Button In Me.Controls
        If TypeOf Button Is MSForms.CommandButton Then
        
            If CByte(Right(Button.Name, 1)) = button_num Then
                Set FindButtonByNumber = Button
                Exit Function
            End If
        
        End If
    Next Button
    
    Err.Raise "2", "FormsBoard", "Button do not exist!"
    
End Function
