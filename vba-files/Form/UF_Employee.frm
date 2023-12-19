VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Employee 
   Caption         =   "Employee Details"
   ClientHeight    =   8595.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17472
   OleObjectBlob   =   "UF_Employee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim ctrlCollection As Collection

Private Sub cbo_DbType_Change()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cbo_DbType_Change"
   
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
   
   Me.db_ID.Value = Date
   Me.cmd_12.Enabled = True 'Enable Save button
   
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub cbo_DEBRIEFtype_Change()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cbo_DEBRIEFtype_Change"
   
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
    
    Me.cmd_12.Enabled = True 'Enable Save button
    Me.date_DEBRIEF = Date
    Me.cmd_12.SetFocus
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub cbo_DEPT_Change()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cbo_DEPT_Change"
   
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
    
    Me.cmd_12.Enabled = True 'Enable Save button
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
End Sub

Private Sub cbo_EligStatus_Change()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cbo_EligStatus_Change"
   
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
   
    If Me.cbo_ELIGstatus.Value = "Interim Secret" Then
        Me.cbo_PSQstatus.Value = "Needs NDA"
    ElseIf Me.cbo_ELIGstatus.Value = "Secret" Then
        If Me.date_NDA = vbNullString Then
            Me.cbo_PSQstatus.Value = "Needs NDA"
        Else
            Me.cbo_PSQstatus.Value = vbNullString
        End If
    End If
    Me.date_PSQdue.Value = vbNullString
    Me.date_ELIG.Value = Date
    Me.cmd_12.Enabled = True 'Enable Save button
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
        
End Sub

Private Sub cbo_EMPstatus_Change()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cbo_EMPstatus_Change"
   
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
   
    If Me.cbo_EMPstatus.Value = "Terminated" Then
        Me.date_TERM.Value = Date
        If Not IsNull(Me.date_ACCESS.Value) Then
            Me.date_REMOVED.Value = Date
            Me.cbo_DEBRIEFtype.SetFocus
        End If
    End If
    
    Me.cmd_12.Enabled = True 'Enable Save button
   
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub cbo_INVtype_Change()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cbo_INVtype_Change"
   
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If

    'Me.cmd_12.Enabled = True 'Enable Save button
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub cbo_NAMEsuffix_Change()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cbo_NAMEsuffix_Change"
   
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
    
    Me.cmd_12.Enabled = True 'Enable Save button
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub cbo_PSQstatus_Change()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cbo_PSQstatus_Change"
   
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
   
    If Me.cbo_PSQstatus.Value = "FSO Review" Then
        Me.date_PSQdue.Value = Date
    ElseIf Me.cbo_PSQstatus.Value = "Paper Version" Then
        Me.date_PSQdue.Value = Date + 13
    ElseIf Me.cbo_PSQstatus.Value = "Corrected Copy" Or Me.cbo_PSQstatus.Value = "Applicant Release" Then
        Me.date_PSQdue.Value = Date + 6
    ElseIf Me.cbo_PSQstatus.Value = "Sent to ISP" Then
        Me.date_PSQdue.Value = vbNullString
    ElseIf Me.cbo_PSQstatus.Value = "PSQ Terminated" Then
        Me.date_PSQdue.Value = vbNullString
        If cbo_INVtype.Value = "Tier 3R" Or cbo_INVtype.Value = "Tier 5R" Then GoTo ExitSub
        Me.cbo_ELIGstatus.Value = "None"
        Me.date_ELIG.Value = Date
    End If
    
    Me.cmd_12.Enabled = True 'Enable Save button
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub cmd_1_Click()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cmd_1_Click"
    
    If Me.txt_NAMElast.Enabled = True Then
        Control_NBF False
    Else
        Control_NBF True
    End If
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub cmd_6_Click()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cmd_6_Click"
   
    CloseUF
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub cmd_7_Click()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cmd_7_Click"
    
    With UF_InvStart
        '.Tag = tagID '<~~ tell the UserForm there's something to bring in so that it'll fill controls from the sheet instead of initializing them
        .show
        '.Tag = "" '<~~ bring Tag property back to its "null" value
    End With

ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub cmd_12_Click()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "cmd_12_Click"
   
    SaveUF_Validate
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub txt_FName_AfterUpdate()

    Me.txt_NAMEfirst.Value = StrConv(Me.txt_NAMEfirst.Value, vbProperCase)
    Name_Full
    
End Sub

Private Sub txt_LName_AfterUpdate()

    Me.txt_NAMElast.Value = StrConv(Me.txt_NAMElast.Value, vbProperCase)
    Name_Full
    
End Sub

Private Sub txt_MName_AfterUpdate()

    Me.txt_NAMEmiddle.Value = StrConv(Me.txt_NAMEmiddle.Value, vbProperCase)
    Name_Full
    
End Sub

Private Sub UserForm_Activate()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "UserForm_Activate"
'   SwitchOff (True) 'Enable VBA_Performance module vbNullString
    
    Load_cbo '<~~ Load comboboxes
    Control_Buttons False '<~~ Hides unused control buttons
    If Me.Tag = vbNullString Then '<~~ if there's no info from Tag property...
'        Control_NBF False '<~~ Enable Change of name and birt fields
        InitializeValues '<~~ ... then Initialize controls values
    Else
        Populate_UF_Values '<~~ ...otherwise fill controls with recordset values
        Control_NBF False '<~~ Disable Change of name and birt fields
        Me.cmd_12.Enabled = False '<~~ Disable save button till change
    End If
    
ExitSub:
'    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub InitializeValues()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "InitializeValues"
'   SwitchOff (True) 'Enable VBA_Performance module vbNullString
   
    With Me
        .Caption = "Add New Employee"
        '.serial.Value = "actText1"
        '.created_on.Value = "actText2"
        '.created_by.Value = "actText3"
        ' and so on...
    End With
    
ExitSub:
'    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Populate_UF_Values()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Populate_UF_Values"
    
    Dim rst As New ADODB.Recordset
    Dim qry As String
    Dim srcString As String
    
    srcString = Me.Tag

    ConnectDatabase
    
    If Not IsNullOrEmpty(srcString) Then
        qry = "SELECT * FROM EmpDatabase WHERE db_ID = " & srcString
    Else
        qry = "SELECT * FROM EmpDatabase WHERE db_ID = " & 0
    End If
    
    rst.Open qry, DBCON, adOpenKeyset, adLockOptimistic
    
    If rst.RecordCount = 0 Then
        rst.AddNew
    End If
   
    SubSkip = True
    
    With Me
    'On Error GoTo ConnErr
    
        Dim fld As ADODB.Field
        Dim ctrl As control
        Dim fldName As String

        For Each fld In rst.Fields
            If Not fld.Name = "info_Tag" Then
                fldName = fld.Name
                Orkin "Field Name: " & fldName
                Set ctrl = .Controls(fldName)
                If Left(ctrl.Name, 5) = "date_" And Not ctrl.Value = vbNullString Then
                    ctrl.Value = Format(CDate(fld.Value), "MM/DD/YYYY")
                ElseIf ctrl.Name = "num_SSN" Then
                    ctrl.Value = Format(fld.Value, "000-00-0000")
                ElseIf ctrl.Name = "num_PHONE" Then
                    ctrl.Value = Format(fld.Value, "(000) 000-0000")
                ElseIf ctrl.Name = "info_Tag" Then
                    Orkin ctrl.Value
                ElseIf IsNull(fld.Value) Then
                    ctrl.Value = Null
                ElseIf Not IsNull(fldName) Then
                    ctrl.Value = fld.Value
                Else
                    GoTo ExitSub
                End If
                Orkin ctrl.Value & " loaded to " & ctrl.Name
            End If
        Next
    End With
    
    Me.Caption = "Employee Details for: " & Me.File_As.Value
    
    rst.Close
    CloseDatabase

ExitSub:
    SubSkip = False
    SubName = CallingSubName
    Exit Sub
    
ConnErr:
    Dim pos As String
    pos = "*** Connection Error! ***"
    pos = pos & vbCrLf & "Fault in: " & fldName
    Orkin pos
    Resume Next
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName & " --" & fldName
    Err.Clear
    CloseDatabase
    GoTo ExitSub
    
End Sub

Function FindValue(vtf As Integer) As Integer
    
StartSub:
   On Error GoTo ExitSub
   
    Dim found As Range
    Dim ws As Worksheet
    
    Set ws = Sheets("Alerts")
    
    'return the value of function
    Set found = ws.Range("D:D").Find(vtf)
    FindValue = found.row
    Orkin "Function FindValue used"
    
ExitSub:
    Exit Function

End Function

Private Sub RecordValues()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "RecordValues"
   
    Dim ws As Worksheet
    Dim rowNum As Integer
    Dim tagValue As Integer
    Set ws = Sheets("Alerts")
    tagValue = Me.Tag
    rowNum = FindValue(tagValue)
    
    SubSkip = True
    
    With ws
    'Name
'        .Cells(rowNum, 1).Value = Me.db_ID.Value
'        .Cells(rowNum, 6).Value = Me.File_As.Value
        .Cells(rowNum, 4).Value = Me.txt_NAMElast.Value
        .Cells(rowNum, 2).Value = Me.txt_NAMEfirst.Value
        .Cells(rowNum, 3).Value = Me.txt_NAMEmiddle.Value
        .Cells(rowNum, 5).Value = Me.cbo_NAMEsuffix.Value
        .Cells(rowNum, 7).Value = Me.num_SSN.Value

    'Birth
        .Cells(rowNum, 8).Value = Me.date_BIRTH.Value
        .Cells(rowNum, 9).Value = Me.txt_bCITY.Value
        .Cells(rowNum, 10).Value = Me.txt_bSTATE.Value

    'Employment
        .Cells(rowNum, 14).Value = Me.date_HIRE.Value
        .Cells(rowNum, 15).Value = Me.date_TERM.Value
        .Cells(rowNum, 16).Value = Me.cbo_DEPT.Value
        .Cells(rowNum, 17).Value = Me.cbo_EMPstatus.Value

    'Contact
        .Cells(rowNum, 11).Value = Me.txt_ADDRESS.Value
        .Cells(rowNum, 12).Value = Me.num_PHONE.Value
        .Cells(rowNum, 13).Value = Me.txt_EMAIL.Value

    'Clearance
    'FP.Value
        .Cells(rowNum, 18).Value = Me.date_FP.Value
        .Cells(rowNum, 19).Value = Me.date_SAC.Value

    'Diss
        .Cells(rowNum, 20).Value = Me.date_DISS.Value
        .Cells(rowNum, 30).Value = Me.date_ACCESS.Value
        .Cells(rowNum, 32).Value = Me.date_REMOVED.Value
        .Cells(rowNum, 31).Value = Me.date_ALTESS.Value
        .Cells(rowNum, 33).Value = Me.cbo_DEBRIEFtype.Value
        .Cells(rowNum, 34).Value = Me.db_ID.Value

    'Eligibility
        .Cells(rowNum, 26).Value = Me.cbo_ELIGstatus.Value
        .Cells(rowNum, 27).Value = Me.date_ELIG.Value
        .Cells(rowNum, 28).Value = Me.date_CE.Value

    'e-QIP
        .Cells(rowNum, 24).Value = Me.cbo_PSQstatus.Value
        .Cells(rowNum, 25).Value = Me.date_PSQdue.Value

    'Inv
        .Cells(rowNum, 23).Value = Me.cbo_INVtype.Value
        .Cells(rowNum, 21).Value = Me.date_INVopen.Value
        .Cells(rowNum, 22).Value = Me.date_INVclose.Value
        .Cells(rowNum, 29).Value = Me.date_NDA.Value
    End With

ExitSub:
    SubSkip = False
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub Load_cbo()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Load_cbo"
    
    SubSkip = True
    With Me
        With cbo_NAMEsuffix
            .AddItem vbNullString
            .AddItem "JR"
            .AddItem "SR"
            .AddItem "I"
            .AddItem "II"
            .AddItem "III"
            .Value = vbNullString
        End With
    
        With cbo_DEPT
            .AddItem "Administration"
            .AddItem "Duplicating"
            .AddItem "Fire"
            .AddItem "Janitorial"
            .AddItem "Security"
            .Value = "Administration"
        End With
    
        With cbo_EMPstatus
            .AddItem "Current"
            .AddItem "Leave"
            .AddItem "Terminated"
            .Value = "Current"
        End With
        
        With cbo_Site
            .AddItem vbNullString
            .AddItem "Radford"
            .Value = "Radford"
        End With
        
        With cbo_INVtype
            .AddItem vbNullString
            .AddItem "NACLC"
            .AddItem "RSI"
            .AddItem "SSBI"
            .AddItem "Tier 3"
            .AddItem "Tier 3R"
            .AddItem "Tier 5"
            .AddItem "Tier 5R"
            .Value = vbNullString
        End With
        
        With .cbo_ELIGstatus
            .AddItem vbNullString
            .AddItem "CL Not Required"
            .AddItem "None"
            .AddItem "SAC (CAC Only)"
            .AddItem "PSQ Initialized"
            .AddItem "PSQ Complete"
            .AddItem "Elig Pending"
            .AddItem "Interim Secret"
            .AddItem "Interim SCI"
            .AddItem "Secret"
            .AddItem "TS/SCI"
            .AddItem "Confendential"
            .Value = vbNullString
        End With
    
        With cbo_PSQstatus
            .AddItem vbNullString
            .AddItem "Paper Version"
            .AddItem "Review Copy"
            .AddItem "Corrected Copy"
            .AddItem "FSO Review"
            .AddItem "Applicant Release"
            .AddItem "Sent to ISP"
            .AddItem "Needs NDA"
            .AddItem "PSQ Stoped"
            .AddItem "PSQ Terminated"
            .Value = vbNullString
        End With
        
        With cbo_DEBRIEFtype
            .AddItem vbNullString
            .AddItem "Admin"
            .AddItem "Person"
            .Value = vbNullString
        End With
    End With
    
ExitSub:
    SubSkip = False
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub
Private Sub SaveChange()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "SaveChange"
   
    Dim rst As New ADODB.Recordset
    Dim qry As String

    ConnectDatabase
    
    If Not IsNullOrEmpty(Me.db_ID.Value) Then
        qry = "SELECT * FROM EmpDatabase WHERE db_ID = " & Me.db_ID.Value
    Else
        qry = "SELECT * FROM EmpDatabase WHERE db_ID = " & 0
    End If
    
    rst.Open qry, DBCON, adOpenKeyset, adLockOptimistic
    
    If rst.RecordCount = 0 Then
        rst.AddNew
    End If
    
    With Me
        On Error GoTo ConnErr
        Dim fld As ADODB.Field
        For Each fld In rst.Fields
        If Not fld.Name = "info_Tag" Then
            Dim fldName As String
            fldName = fld.Name
            Dim ctrl As control
            Set ctrl = .Controls(fldName)
            If Not fldName = "db_ID" Then
                If Left(ctrl.Name, 5) = "date_" And Not ctrl.Value = vbNullString Then
                    fld.Value = Format(CDate(ctrl.Value), "MM/DD/YYYY")
                ElseIf IsNullOrEmpty(ctrl.Value) Then
                    fld.Value = Null
                Else
                    fld.Value = ctrl.Value
                End If
                Orkin "Saved " & ctrl.Value & " To " & fld.Name
            End If
        End If
        Next
    End With
    
On Error GoTo ErrCtrl
       
    rst.Update
    rst.Close
    CloseDatabase
       
ExitSub:
    SubName = CallingSubName
    CloseUF
    Exit Sub
    
ConnErr:
    Dim pos As String
    pos = "*** Connection Error! ***"
    pos = pos & vbCrLf & "Fault in: " & fld.Name
    Orkin pos
    Resume Next
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    CloseDatabase
    GoTo ExitSub
    
End Sub
Private Sub CloseUF()

StartSub:
   On Error GoTo ExitSub
   
   Unload Me
   
ExitSub:
    Exit Sub
    
End Sub

Private Sub Control_Buttons(xCtrl As Boolean)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Control_Buttons"
    '   SwitchOff (True) 'Enable VBA_Performance module
    
    Dim ctrl As control
    Dim ubCtrl As Boolean

    With Me
    'Tag used buttons
        .cmd_1.Tag = "cmdbtn_NBF" 'Edit name & bith fields
        .cmd_6.Tag = "cmdbtn_cancel" 'Cancel button
        .cmd_7.Tag = "cmdbtn_start_inv" 'Start inv button
        .cmd_12.Tag = "cmdbtn_save" 'Save button
        
    'Tag unused buttons
        .cmd_2.Tag = "unused_btn2" 'unused button
        .cmd_3.Tag = "unused_btn3" 'unused button
        .cmd_4.Tag = "unused_btn4" 'unused button
        .cmd_5.Tag = "unused_btn5" 'unused button
        .cmd_8.Tag = "unused_btn8" 'unused button
        .cmd_9.Tag = "unused_btn9" 'unused button
        .cmd_10.Tag = "unused_btn10" 'unused button
        .cmd_11.Tag = "unused_btn11" 'unused button
        
        For Each ctrl In frm_BTNS.Controls
            ubCtrl = xCtrl
            Select Case TypeName(ctrl)
                Case "CommandButton"
                If Left(ctrl.Tag, 7) = "cmdbtn_" Then ubCtrl = True 'Enable used buttons
                ctrl.Enabled = ubCtrl
                ctrl.Visible = ubCtrl
            End Select
        Next
        
        .cmd_12.Enabled = False '<~~ Disable save button till change
    End With

ExitSub:
'    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Control_NBF(xCtrl As Boolean)

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "Control_NBF"
'   SwitchOff (True) 'Enable VBA_Performance module vbNullString

    Dim ctrl As control
    
'Name
    For Each ctrl In frm_Name.Controls
        Select Case TypeName(ctrl)
        Case "TextBox", "ComboBox"
            ctrl.Enabled = xCtrl
        End Select
    Next
    
'Birth
    For Each ctrl In frm_BIRTH.Controls
        Select Case TypeName(ctrl)
        Case "TextBox", "ComboBox"
            ctrl.Enabled = xCtrl
        End Select
    Next
    
' Always disable ID field
    Me.db_ID.Enabled = False '<~~ Disable user change of ID Field
    GoTo ExitSub
    
ExitSub:
'    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub
Private Sub UserForm_Initialize()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "UserForm_Initialize"
'   SwitchOff (True) 'Enable VBA_Performance module
    
    If Application.Left > 1000 Then
        With Me
            .Top = (Application.UsableHeight / 2) + (Me.Height / 2)
            .Left = (1.5 * (Application.UsableWidth)) - (Me.Width / 2)
            .StartUpPosition = 1
        End With
    Else
        With Me
            .Top = (Application.UsableHeight / 2) + (Me.Height / 2)
            .Left = (Application.UsableWidth / 2) - (Me.Width / 2)
            .StartUpPosition = 1
        End With
    End If
    
    Enable_Save '<~~ ... enable save button if there is changes

ExitSub:
'    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub
Private Sub SaveUF_Validate()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "SaveUF_Validate"

    SubSkip = True
    
    Dim ufContinue As Integer
    Dim dateStamp As Date

    dateStamp = Date

    With Me
    'Name
        .txt_NAMEfirst.Value = StrConv(.txt_NAMEfirst.Value, vbProperCase)
        .txt_NAMEmiddle.Value = StrConv(.txt_NAMEmiddle.Value, vbProperCase)
        .txt_NAMElast.Value = StrConv(.txt_NAMElast.Value, vbProperCase)

    'Full Name
        Name_Full

    'SSN
        If IsNullOrEmpty(.num_SSN.Value) Then
            ufContinue = MsgBox("SSN required." & vbNewLine & "Do you want to continue without?", vbQuestion + vbYesNo + vbDefaultButton2, "Required Information")
            If ufContinue = vbYes Then
                .num_SSN.Value = vbNullString
            Else
                GoTo valCancel
            End If
        Else
            .num_SSN.Value = CleanSsnNumber(.num_SSN.Value)
        End If

    'Birth
        If IsNullOrEmpty(.date_BIRTH.Value) Then
            ufContinue = MsgBox("Date of birth required." & vbNewLine & "Do you want to continue without?", vbQuestion + vbYesNo + vbDefaultButton2, "Required Information")
            If ufContinue = vbYes Then
                .date_BIRTH.Value = Null
            Else
                GoTo valCancel '<~~~ Return to UserForm
            End If
        ElseIf IsDate(.date_BIRTH.Value) = False Then
            MsgBox "Please enter the date of birth" & vbNewLine & "(MM/DD/YYYY)", vbCritical, "Required Information"
            GoTo valCancel '<~~~ Return to UserForm
        ElseIf Me.date_BIRTH.Value >= dateStamp Then
            MsgBox "Date of birth cannot be equal to or greater than today", vbCritical, "Required Information"
            GoTo valCancel '<~~~ Return to UserForm
        Else
            .date_BIRTH.Value = Format(CDate(.date_BIRTH.Value), "MM/DD/YYYY")
        End If
        
'        .txt_bCity.Value = .txt_bCity.Value
'        .txt_bState.Value = .txt_bState.Value

    'Contact
'        .txt_Address.Value = .txt_Address.Value
        If IsNullOrEmpty(.num_PHONE.Value) Then
            ufContinue = MsgBox("Phone # required." & vbNewLine & "Do you want to continue without?", vbQuestion + vbYesNo + vbDefaultButton2, "Required Information")
            If ufContinue = vbYes Then
                .num_PHONE.Value = Null
            Else
                GoTo valCancel '<~~~ Return to UserForm
            End If
        Else
           .num_PHONE.Value = CleanPhoneNumber(.num_PHONE.Value)
        End If
        If IsNullOrEmpty(.txt_EMAIL.Value) Then
            ufContinue = MsgBox("Email required." & vbNewLine & "Do you want to continue without?", vbQuestion + vbYesNo + vbDefaultButton2, "Required Information")
            If ufContinue = vbYes Then
                .txt_EMAIL.Value = Null
            Else
                GoTo valCancel '<~~~ Return to UserForm
            End If
'        Else
'            .txt_EMAIL.Value = .txt_EMAIL.Value
        End If

    'Employment
        If IsNullOrEmpty(.date_HIRE.Value) Then
            ufContinue = MsgBox("Hire date required." & vbNewLine & "Do you want to continue without?", vbQuestion + vbYesNo + vbDefaultButton2, "Required Information")
            If ufContinue = vbYes Then
                .date_HIRE.Value = Null
            Else
                GoTo valCancel '<~~~ Return to UserForm
            End If
        Else
            .date_HIRE.Value = Format(CDate(.date_HIRE.Value), "MM/DD/YYYY")
        End If

    'FP.Value
        If IsNullOrEmpty(.date_FP.Value) Then
            .date_FP.Value = Null
        Else
            .date_FP.Value = Format(CDate(.date_FP.Value), "MM/DD/YYYY")
        End If

        If IsNullOrEmpty(.date_SAC.Value) Then
            .date_SAC.Value = Null
        Else
            .date_SAC.Value = Format(CDate(.date_SAC.Value), "MM/DD/YYYY")
        End If

    'Diss
        If IsNullOrEmpty(.date_DISS.Value) Then
            .date_DISS.Value = Null
        Else
            .date_DISS.Value = Format(CDate(.date_DISS.Value), "MM/DD/YYYY")
        End If

    'Inv
        If IsNullOrEmpty(.date_INVopen.Value) Then
            .date_INVopen.Value = Null
        Else
            .date_INVopen.Value = Format(CDate(.date_INVopen.Value), "MM/DD/YYYY")
        End If
        
        If IsNullOrEmpty(.date_INVclose.Value) Then
            .date_INVclose.Value = Null
        Else
            .date_INVclose.Value = Format(CDate(.date_INVclose.Value), "MM/DD/YYYY")
        End If

'        cbo_invType.Value = .cbo_invType.Value

    'e-QIP
'        .cbo_PSQstatus.Value = .cbo_PSQstatus.Value

        If IsNullOrEmpty(.date_PSQdue.Value) Then
            .date_PSQdue.Value = Null
        Else
            .date_PSQdue.Value = Format(CDate(.date_PSQdue.Value), "MM/DD/YYYY")
        End If

    'Eligibility
        If IsNullOrEmpty(.date_ELIG.Value) Then
            .date_ELIG.Value = Null
        Else
            .date_ELIG.Value = Format(CDate(.date_ELIG.Value), "MM/DD/YYYY")
        End If

        If IsNullOrEmpty(.date_CE.Value) Then
            .date_CE.Value = Null
        Else
            .date_CE.Value = Format(CDate(.date_CE.Value), "MM/DD/YYYY")
        End If

    'NDA
        If IsNullOrEmpty(.date_NDA.Value) Then
            .date_NDA.Value = Null
        Else
            .date_NDA.Value = Format(CDate(.date_NDA.Value), "MM/DD/YYYY")
        End If

    'Access
        If IsNullOrEmpty(.date_ACCESS.Value) Then
            .date_ACCESS.Value = Null
        Else
            .date_ACCESS.Value = Format(CDate(.date_ACCESS.Value), "MM/DD/YYYY")
        End If

        If IsNullOrEmpty(.date_REMOVED.Value) Then
            .date_REMOVED.Value = Null
        Else
            .date_REMOVED.Value = Format(CDate(.date_REMOVED.Value), "MM/DD/YYYY")
        End If
    
        If IsNullOrEmpty(.date_ALTESS.Value) Then
            .date_ALTESS.Value = Null
        Else
            .date_ALTESS.Value = Format(CDate(.date_ALTESS.Value), "MM/DD/YYYY")
        End If

    'Debrief
'        .cbo_DEBRIEFtype.Value = .cbo_DEBRIEFtype.Value

    'Tag
'        If IsNullOrEmpty(.date_FP.Value) And cbo_ELIGstatus = "None" Then
'            txt_Tag.Value = "Need Prints"
'
'        txt_Tag.Value = "PR Due"
    
    End With
    
    SaveChange '<~~~ Save changes

ExitSub:
    SubName = CallingSubName
    SubSkip = False
    Exit Sub
    
valCancel:
    ufContinue = MsgBox("Save Cancled!" & vbNewLine & "Do you want to edit employee?", vbQuestion + vbYesNo + vbDefaultButton1, "Validation Failure")
            If ufContinue = vbYes Then
                GoTo ExitSub
            Else
                SubName = CallingSubName
                SubSkip = False
                Unload Me
            End If
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub Enable_Save()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "Enable_Save"
'   SwitchOff (True) 'Enable VBA_Performance module

    Dim ctrl As MSForms.control
    Dim obj As clsTextBox
    
    Set ctrlCollection = New Collection
        For Each ctrl In Me.Controls
            If TypeOf ctrl Is MSForms.TextBox Then
                Set obj = New clsTextBox
                Set obj.tbControl = ctrl
                ctrlCollection.Add obj
            ElseIf TypeOf ctrl Is MSForms.ComboBox Then
                Set obj = New clsTextBox
                Set obj.cbControl = ctrl
                ctrlCollection.Add obj
            End If
        Next ctrl
    Set obj = Nothing
    
ExitSub:
'    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub


Private Sub Name_Full()

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "Name_Full"
   
    With Me
        If IsNullOrEmpty(.cbo_NAMEsuffix.Value) Then
            .File_As.Value = UCase(.txt_NAMElast.Value) & ", " & .txt_NAMEfirst.Value & " " & Left(.txt_NAMEmiddle.Value, 1)
        Else
            .File_As.Value = UCase(.txt_NAMElast.Value) & " " & Me.cbo_NAMEsuffix.Value & ", " & .txt_NAMEfirst.Value & " " & Left(.txt_NAMEmiddle.Value, 1)
        End If
    End With
   
ExitSub:
'    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Function CheckPhoneNumber(PhoneNumber As String) As Boolean
    
    CheckPhoneNumber = False
    
    Dim PhoneNumberLength As Long
    PhoneNumberLength = Len(PhoneNumber)
    
    If PhoneNumberLength = 10 Then CheckPhoneNumber = True
        
End Function

Private Function CleanPhoneNumber(PhoneNumber As String) As String
    
    If Not PhoneNumber = vbNullString Then
        Dim i As Long
        For i = 1 To Len(PhoneNumber)
            If Asc(Mid(PhoneNumber, i, 1)) >= Asc("0") And Asc(Mid(PhoneNumber, i, 1)) <= Asc("9") Then
            Dim RetainNumber As String
            RetainNumber = RetainNumber + Mid(PhoneNumber, i, 1)
        End If
        Next
    Else
        Exit Function
    End If
    
    If CheckPhoneNumber(RetainNumber) Then CleanPhoneNumber = RetainNumber
End Function


Private Function CheckSsn(SsNumber As String) As Boolean
    
    CheckSsn = False
    
    Dim CheckSsnLength As Long
    CheckSsnLength = Len(SsNumber)
    
    If CheckSsnLength = 9 Then CheckSsn = True
        
End Function

Private Function CleanSsnNumber(SsNumber As String) As String
    
    If Not SsNumber = vbNullString Then
        Dim i As Long
        For i = 1 To Len(SsNumber)
            If Asc(Mid(SsNumber, i, 1)) >= Asc("0") And Asc(Mid(SsNumber, i, 1)) <= Asc("9") Then
            Dim RetainNumber As String
            RetainNumber = RetainNumber + Mid(SsNumber, i, 1)
        End If
        Next
    Else
        Exit Function
    End If
    
    If CheckSsn(RetainNumber) Then CleanSsnNumber = RetainNumber
End Function
