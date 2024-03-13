VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub InitializeApplication()
' Entry point of the Application.

    Dim settings As New Scripting.Dictionary
    
    Application.WindowState = xlMaximized
    Application.Visible = False
    
    Set settings = Me.Sheets(1).GatherSettings()
    Main.RunGame settings
    
End Sub

Public Sub ExitWorkbook()
' Procedure responsible for exitting the workbook.

    Application.DisplayAlerts = False
    Application.Quit
    
End Sub

Private Sub Workbook_Activate()
' Initializes default settings in sheet whenever it is opened.
    
    Me.Sheets(1).WorksheetInitialize

End Sub

Private Sub Workbook_BeforeClose(cancel As Boolean)
' Event handler for exiting the workbook.

    ExitWorkbook
    
End Sub

