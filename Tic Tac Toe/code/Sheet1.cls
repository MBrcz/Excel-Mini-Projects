VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ------------------------------------
' ----- PRIVATE ENUMS -----
' ------------------------------------

Private Enum ColumnPlayer
' Represents column that are bound to the player.
    columnplayer1 = 2
    columnplayer2 = 3
End Enum

Private Enum RowConstants
' Represents rows where setting is stored.
    Name = 3
    Icon = 4
    PlayerType = 5
    ComputerDifficulity = 6
    IconApperance = 7
End Enum

' --------------------------------------
' ------- WORKSHEET ----------
' --------------------------------------
Public Sub WorksheetInitialize()
    
    ChangeCheckBox Me.Pl1CheckBox, Me.Pl2Checkbox, Me.Pl1ComboDiff, Me.Pl1Textbox, 1
    ChangeCheckBox Me.Pl2Checkbox, Me.Pl1CheckBox, Me.Pl2ComboDiff, Me.Pl2TextBox, 2
    
    ChangeComboBox ColumnPlayer.columnplayer1
    ChangeComboBox ColumnPlayer.columnplayer2

End Sub

' ----------------------------------------
' --------- CONTROLS ------------
' ----------------------------------------

Private Sub Pl2Checkbox_Change()
    
     ChangeCheckBox Me.Pl2Checkbox, Me.Pl1CheckBox, Me.Pl2ComboDiff, Me.Pl2TextBox, 2
    
End Sub

Private Sub Pl1CheckBox_change()
    
    ChangeCheckBox Me.Pl1CheckBox, Me.Pl2Checkbox, Me.Pl1ComboDiff, Me.Pl1Textbox, 1

End Sub

Private Sub Pl1ComboBox_Change()

    ChangeComboBox ColumnPlayer.columnplayer1

End Sub

Private Sub Pl2ComboBox_Change()

    ChangeComboBox ColumnPlayer.columnplayer2

End Sub

' --------------------------------------
' ------------- Settings ------------
' --------------------------------------

Public Function GatherSettings() As Scripting.Dictionary
' Gets all setting from the first sheet.

    Dim local_settings As New Scripting.Dictionary
    
    With local_settings
        .Add Enums.TranslateSettings(Player1Name), Me.Pl1Textbox.value
        .Add Enums.TranslateSettings(Player2Name), Me.Pl2TextBox.value

        .Add Enums.TranslateSettings(Player1Icon), Me.Pl1ComboBox.value
        .Add Enums.TranslateSettings(Player2Icon), Me.Pl2ComboBox.value

        .Add Enums.TranslateSettings(Player1Number), eBoardMove.Player1Move
        .Add Enums.TranslateSettings(Player2Number), eBoardMove.Player2Move

        .Add Enums.TranslateSettings(Player1Type), AssignPlayer(Me.Pl1CheckBox)
        .Add Enums.TranslateSettings(Player2Type), AssignPlayer(Me.Pl2Checkbox)

        .Add Enums.TranslateSettings(player1ComputerType), Me.Pl1ComboDiff.value
        .Add Enums.TranslateSettings(player2ComputerType), Me.Pl2ComboDiff.value
    End With

    Set GatherSettings = local_settings

End Function

Private Function AssignPlayer(CheckBox As MSForms.CheckBox) As String
' Assigns the suitable player to the checkboxes.
    
    If CheckBox.value = True Then
        AssignPlayer = Enums.TranslatePlayerTypes(AI)
    
    ElseIf CheckBox.value = False Then
        AssignPlayer = Enums.TranslatePlayerTypes(HUMAN)
    End If

End Function

' --------------------------------------
' -------- CHECKBOXES --------
' --------------------------------------

Private Function ChangeCheckBox(FirstCheckBox As MSForms.CheckBox, SecondCheckBox As MSForms.CheckBox, _
                                                      DiffWin As MSForms.ComboBox, TextBox As MSForms.TextBox, diff_num As Byte)
    ' Handles the change of the is_computer checkbox.

    If FirstCheckBox.value = True Then
        SecondCheckBox.value = False
        DiffWin.Enabled = True
        TextBox.value = Enums.TranslatePlayerTypes(AI)
        
    ElseIf FirstCheckBox.value = False Then
        DiffWin.Enabled = False
        TextBox.value = Enums.TranslatePlayerTypes(HUMAN) & " " & CStr(diff_num)
    End If
    
    AddComputerDifficulities DiffWin

End Function

' ---------------------------------------
' ------ COMBOBOXES ----------
' ---------------------------------------

Private Function ChangeComboBox(playernum As Byte)
' Handles changing the combox of the Icon type.
    
    'SwitchComboBoxStyles Me.Pl1ComboBox, Me.Pl2ComboBox
    
    If playernum = ColumnPlayer.columnplayer1 Then
        PopulateComboBox Me.Pl1ComboBox, Me.Pl2ComboBox
    ElseIf playernum = ColumnPlayer.columnplayer2 Then
        PopulateComboBox Me.Pl2ComboBox, Me.Pl1ComboBox
    End If
    
    UpdateImages

End Function

Private Function AddComputerDifficulities(ComboBox As MSForms.ComboBox)
' Fills elements to the difficulity combobox'es.
    
    With ComboBox
        .List = Array(Enums.TranslateComputerMode(mixed), Enums.TranslateComputerMode(random), Enums.TranslateComputerMode(UNBEATABLE))
    End With

End Function

Private Function PopulateComboBox(FirstCombobox As MSForms.ComboBox, SecondCombobox As MSForms.ComboBox)
' Fills elements to choose from the combo box.

    Dim selected_value As String
    Dim image_names() As Variant
    Dim num As Byte
    
    image_names = utils.GetAllFileNamesFromDir(utils.GetPathToImagesFile)
    
    FirstCombobox.List = Array()
    SecondCombobox.List = Array()
    
     For num = LBound(image_names) To UBound(image_names)
        If image_names(num) <> SecondCombobox.value Then
            FirstCombobox.AddItem image_names(num)
        End If
        
        If image_names(num) <> FirstCombobox.value Then
            SecondCombobox.AddItem image_names(num)
        End If
    Next num
    
End Function

Private Sub SwitchComboBoxStyles(ParamArray ComboBoxes() As Variant)
' Turns on or off variable ammount of ComboBoxes.
    
    Dim cb As MSForms.Control
    Dim i As Integer
    
    For i = LBound(ComboBoxes) To UBound(ComboBoxes)
        If TypeOf ComboBoxes(i) Is MSForms.Control Then
            Set cb = ComboBoxes(i)
            If TypeName(cb) = "ComboBox" Then
                If cb.Style = fmStyleDropDownCombo Then
                    cb.Style = fmStyleDropDownList
                ElseIf cb.Style = fmStyleDropDownList Then
                    cb.Style = fmStyleDropDownCombo
                End If
            End If
        End If
    Next i
    
End Sub

' ----------------------------------------
' --------- IMAGES ------------------
' ----------------------------------------

Private Function UpdateImages()
' Handles dealing with the showing image in the Worksheet.
    
    Dim image_dir As String
    
    image_dir = utils.GetPathToImagesFile()
    
    RemoveAllImagesWorksheet
    PlaceImageInWorksheet image_dir & Me.Pl1ComboBox.value, ColumnPlayer.columnplayer1
    PlaceImageInWorksheet image_dir & Me.Pl2ComboBox.value, ColumnPlayer.columnplayer2

End Function

Private Function RemoveAllImagesWorksheet()
    ' Removes all images from the worksheet.
    Dim shape As shape
    
    For Each shape In Me.Shapes
        If shape.Type = msoPicture Then
            shape.Delete
        End If
    Next shape
End Function

Private Function PlaceImageInWorksheet(image_path As String, column As Byte)
' Places an image at the certain row and column in the worksheet.
    
    Dim TargetCell As Range
    Dim imgShape As shape
    
    Set TargetCell = Cells(IconApperance, column)
    
    ' Create a Shape object by adding the picture
    Set imgShape = Me.Shapes.AddPicture(Filename:=image_path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                        Left:=1, Top:=1, Width:=-1, Height:=-1)
    
    TeleportImageToCell TargetCell, imgShape
    
End Function

Private Function TeleportImageToCell(TargetCell As Range, ImageObject As shape)
' Places an image exactly in cell and makes it size equal's to the cell's

    With ImageObject
        .LockAspectRatio = msoFalse
        .Width = TargetCell.Width
        .Height = TargetCell.Height
        .Left = TargetCell.Left
        .Top = TargetCell.Top
    End With
    
End Function
