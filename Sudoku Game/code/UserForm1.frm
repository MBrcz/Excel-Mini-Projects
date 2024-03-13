VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   13065
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11130
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ------------------------------------------------------------------------------------------------------------------------------------------
' This user form module contains all handlers and events that are related to the initialization of the Board.
' ------------------------------------------------------------------------------------------------------------------------------------------

' property
Private BoardCells() As evCellBinder ' an array of the evCellBinder objects. See more: evCellBinder

Public Function RestartBoard()
    ' Function that resets the whole board and it's logic.
    
    ' Accepts:
    '    None
    ' Returns:
    '   None
    
    Me.Caption = "Sudoku Game"
   
    SetupTheBoard
    mController.loadController BoardCells, Me

End Function

' -------------------------------------------------------------------
' ---------- FORMS EVENT HANDLERS -----------------
' -------------------------------------------------------------------

Private Sub UserForm_Initialize()
    ' Handles the initialization of the UserForm object.
   
    RestartBoard

End Sub

Private Sub Restart_Button_Click()
    ' Event handler for restartButton click

    RestartBoard

End Sub

' --------------------------------------------------------------------
' ----------------- INTIALIZATION -----------------------------
' --------------------------------------------------------------------

Private Function SetupTheBoard()
    ' Setups the board from the GUI point of view in the application.
    ' It creates a BoardCell() 2D array with evCellBinder objects (which represents cells), formats them and assings event to them.
    
    ' Accepts:
    '   None
    '   Return:
    '   None

    Dim subFrame As Object
    Dim mainFrame As frame
    Dim control As Object
    
    ReDim BoardCells(SudokuConstans.FirstRow To SudokuConstans.LastRow, SudokuConstans.FirstRow To SudokuConstans.LastRow)
    
    Set mainFrame = Me.controls("Frame1")
    
    For Each subFrame In mainFrame.controls
        If TypeName(subFrame) = "Frame" Then
            subFrame.Caption = ""
            For Each control In subFrame.controls
                If TypeName(control) = "TextBox" And Not InStr(control.name, "TextBox") <> 0 Then
                    AssignEventsToControl control
                End If
            Next control
        End If
    Next subFrame

End Function

Private Function AssignEventsToControl(cellControl As control)
    ' Assigns the control called cell to the CellBinder object that stores the event callback for each cell
    
    ' Accepts:
    '   cellControl [control that represents cell object]
    ' Returns:
    '   None
        
    Dim position() As Variant
    Dim num As Integer
    Dim cellBinder As evCellBinder
    
    Set cellBinder = New evCellBinder
    num = utils.ExtractNumbersFromString(cellControl.name)
    position = utils.SplitTheNumberIntoPositions(num)

    BindCellObject cellBinder, cellControl, position
    FormatTextBox cellBinder.cell
    Set BoardCells(position(0), position(1)) = cellBinder

End Function
Private Function BindCellObject(cellBinder As evCellBinder, cellControl As control, position As Variant)
    ' Binds cellBinder object and assigns attributes towards it
    
    ' Accepts:
    '   cellBinder [evCellBidner object that will be bound]
    ' Returns:
    '   None
    
    With cellBinder
        Set .cell = cellControl ' textbox
        .cellColumn = position(1) 'xPos in the board
        .cellRow = position(0) ' yPos in the board
        .cellChange = False ' Do not call self _change event durning init
    End With

End Function

Private Function FormatTextBox(ByRef cell As MSForms.textbox)
    ' Formats the cell object as in lower
    
    ' Accepts:
    '   cell [textbox object that will be formatted]
    ' Returns:
    '   None
    
    With cell
        .Font.Size = 30
        .Font.Bold = True
        .BackColor = GetRGBColor(SudokuColors.white)
        .Font.Italic = True
    End With
    
End Function

' ----------------------------------------------------------------------------------------------------------------------------
' ------------------------------ DOWNSIZING OF GUI --------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------------------




