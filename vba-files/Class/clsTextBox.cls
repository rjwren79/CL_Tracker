VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents MyTextBox As MSForms.TextBox
Attribute MyTextBox.VB_VarHelpID = -1
Private WithEvents MyComboBox As MSForms.ComboBox
Attribute MyComboBox.VB_VarHelpID = -1

Public Property Set tbControl(tb As MSForms.TextBox)
    
    Set MyTextBox = tb

End Property

Public Property Set cbControl(cb As MSForms.ComboBox)
    
    Set MyComboBox = cb

End Property

Private Sub MyTextBox_Change()
On Error GoTo ExitSub

StartSub:
    Call Save_btn_Enable
    
ExitSub:
    Exit Sub
End Sub

Private Sub MyComboBox_Change()
On Error GoTo ExitSub

StartSub:
    Call Save_btn_Enable
    
ExitSub:
    Exit Sub
End Sub

Private Sub Save_btn_Enable()
    
StartSub:
    On Error GoTo ExitSub

    UF_Employee.cmd_12.Enabled = True 'Enable Save button
    'AutoCalc() 'call your AutoCalc sub / function whenever textbox changes
    
ExitSub:
    Exit Sub
End Sub
