VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_InvStart 
   Caption         =   "Start Investigation"
   ClientHeight    =   3012
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "UF_InvStart.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UF_InvStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub cmd_Cancel_Click()

StartSub:
   On Error GoTo ExitSub
   
    CloseUF
    
ExitSub:
    Exit Sub
    
End Sub

Private Sub cmd_Save_Click()

StartSub:
   On Error GoTo ExitSub
   
    SaveChange
    
ExitSub:
    Exit Sub
    
End Sub

Private Sub UserForm_Initialize()

StartSub:
   On Error GoTo ExitSub
     
    Load_cbo
    Me.txt_dAddSMO.Value = Date
    If Not IsNullOrEmpty(UF_Employee.date_DISS.Value) Then _
        Me.txt_dAddSMO.Value = UF_Employee.date_DISS.Value
    
ExitSub:
    Exit Sub
    
End Sub

Private Sub SaveChange()

StartSub:
   On Error GoTo ExitSub
   
    With UF_Employee
        If Me.cbo_INVtype.Value = "Paper Version" Then
            .cbo_PSQstatus.Value = "Paper Version"
            .date_PSQdue.Value = Date + 6
            
        Else
            .cbo_INVtype.Value = Me.cbo_INVtype.Value
            .date_DISS.Value = Me.txt_dAddSMO.Value
            .date_INVopen.Value = Date
            .cbo_ELIGstatus.Value = "PSQ Initialized"
            .date_ELIG.Value = Date
            .cbo_PSQstatus.Value = "Review Copy"
            .date_PSQdue.Value = Date + 6
        End If
        
    End With
    
    CloseUF
   
ExitSub:
    Exit Sub
    
End Sub

Private Sub CloseUF()

StartSub:
   On Error GoTo ExitSub
   
   Unload Me
   
ExitSub:
    Exit Sub
    
End Sub

Private Sub Load_cbo()

StartSub:
   On Error GoTo ExitSub
   
    With Me.cbo_INVtype
        .AddItem vbNullString
        .AddItem "Paper Version"
        .AddItem "Confendential"
        .AddItem "Tier 3"
        .AddItem "Tier 3R"
        .AddItem "Tier 5"
        .AddItem "Tier 5R"
    End With
    
ExitSub:
    Exit Sub

End Sub

