Private Sub Workbook_BeforeClose(Cancel As Boolean)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Workbook_BeforeClose"
    SwitchOff (True) 'On 'Enable VBA_Performance module

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
    SwitchOff (True) 'On 'Disable VBA_Performance module
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
    SwitchOff (True) 'On 'Enable VBA_Performance module
    Dim awPath As String
    
    
    'UIHide
    terminix True 'Write debug to immediate window
    
    SheetPrep 'Format sheets, named ranges, and hide developer columns
    
    SubSkip = False
    
ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub
