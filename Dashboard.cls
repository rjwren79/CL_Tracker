
Private Sub Worksheet_Activate()
On Error GoTo ExitSub

    'Application.ScreenUpdating = False
    
    With Worksheets("Dashboard")
        'Cells.Interior.Color = rgb(75, 0, 75)
        GoTo ExitSub
    End With

ExitSub:

   ' Application.ScreenUpdating = True
    Exit Sub

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
On Error GoTo ExitSub

ExitSub:

    'Application.ScreenUpdating = True
    Exit Sub

End Sub
