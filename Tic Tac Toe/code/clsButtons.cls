VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ButtonObject As MSForms.CommandButton
Attribute ButtonObject.VB_VarHelpID = -1

Public Function PlaceImage(image_name As String)
' Function responsible for setting an image in the button.
        
    Dim img As MSForms.image
    Dim ui As UserForm
    Dim image_path As String
    
    Set ui = ButtonObject.Parent
    Set img = ui.Controls.Add("Forms.Image.1", , True)
    image_path = utils.GetPathToImagesFile() & image_name
    
    With img
        .Left = ButtonObject.Left
        .Top = ButtonObject.Top
        .Width = ButtonObject.Width
        .Height = ButtonObject.Height
        .PictureSizeMode = fmPictureSizeModeStretch
        .Picture = LoadPicture(image_path)
    End With
    

End Function

Private Sub ButtonObject_Click()
' Function responsible for handling clicking an button object.

    Dim num As Byte
    num = CByte(Right(ButtonObject.Name, 1))
    
    Main.UpdateContent num

End Sub
