VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Workbook_BeforeClose"
    SwitchOff (True) 'Enable VBA_Performance module

    UIShow
    
    terminix False
    Application.Run "Alerts.ShowCol", False
    Application.Run "Alerts.clear_sh"
    Application.StatusBar = "Done"
    Application.StatusBar = False
    ThisWorkbook.Sheets("Dashboard").Select
    Show_Window True
    ActiveWorkbook.Save

ExitSub:
    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Workbook_Open()

StartSub:
    On Error GoTo ErrCtrl
    SubName = "Workbook_Open"
    SwitchOff (True) 'Enable VBA_Performance module
    Dim awPath As String
    
    
    'UIHide
    terminix True 'Write debug to immediate window
    
    'Sheet_Prep True 'Hide developer columns
    
    SubSkip = False
    
ExitSub:
    SwitchOff (False) 'Disable VBA_Performance module
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

