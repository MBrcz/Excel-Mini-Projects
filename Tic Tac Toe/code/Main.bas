Attribute VB_Name = "Main"
Private AppController As clsController

Private Property Set SetController(value As clsController)
    
    Set AppController = value

End Property

Public Property Get GetController()

    Set GetController = AppController
    
End Property

Public Function RestartGame()
' Restarts the game.

    AppController.RestartGame

End Function

Public Function ExitGame()
' Handles exititng the game object.

    AppController.ExitGame
    
    Set AppController = Nothing
    Application.Visible = True
    
End Function
Public Function UpdateContent(button_num As Byte)
' Handles interaction with button and passes it to the controller object.

    AppController.MakeMove button_num
    
End Function

Public Function RunGame(ByRef settings As Scripting.Dictionary)
' Factory Method

    Dim AppController As New clsController
    
    Set SetController = AppController
    AppController.InitializeGame settings
    
End Function

