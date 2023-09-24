Option Explicit
Private ShwAllColumns As Boolean
Public lstboxColCount As Integer
Public lstboxColWidthStr As String
Dim rstCount As Integer

Private Sub Search_cbo_Filter_AddItem()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Search_cbo_Filter_AddItem"
    SubSkip = True
    With Search_cbo_Filter
        .Clear
        .AddItem "None"
        .AddItem "Alerts"
        .AddItem "Department"
        .AddItem "Employment Status"
        .AddItem "Current Eligibility"
        .AddItem "INV Status"
        .Value = "Alerts"
    End With

ExitSub:
    SubName = CallingSubName
    SubSkip = False
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Search_cbo_Status_AddItem()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Search_cbo_Status_AddItem"
    SubSkip = True
    With Search_cbo_Status
        .Enabled = True
        .Clear
        .AddItem "All"
        Select Case Search_cbo_Filter.Value
        Case "Department"
            .AddItem "Administration"
            .AddItem "Duplicating"
            .AddItem "Fire"
            .AddItem "Janitorial"
            .AddItem "Security"
        Case "Alerts"
            .AddItem "PSQ Over Due"
            .AddItem "Fingerprints Needed"
            .AddItem "Needs Review"
            .AddItem "Ready for Release"
            .AddItem "Needs NDA"
            .AddItem "PSQ Terminated"
            .AddItem "Eligibility Pending"
            .AddItem "PR Due"
        Case "Employment Status"
            .AddItem "Current"
            .AddItem "Leave"
            .AddItem "Terminated"
        Case "Current Eligibility"
            .AddItem "None"
            .AddItem "SAC (CAC Only)"
            .AddItem "PSQ Initialized"
            .AddItem "Elig Pending"
            .AddItem "Interim Secret"
            .AddItem "Secret"
            .AddItem "Interim SCI"
            .AddItem "TS/SCI"
            .AddItem "Confendential"
        Case "INV Status"
            .AddItem "Paper Version"
            .AddItem "Review Copy"
            .AddItem "Corrected Copy"
            .AddItem "FSO Review"
            .AddItem "Applicant Release"
            .AddItem "Sent to ISP"
            .AddItem "Sign NDA"
            .AddItem "PSQ Stoped"
            .AddItem "PSQ Terminated"
        Case Else
            .Clear
            .Value = vbNullString
            .Enabled = False
            GoTo ExitSub
        End Select
        .Value = "All"
    End With
    
ExitSub:
    SubName = CallingSubName
    SubSkip = False
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub
    
Private Sub Search_cbo_Sort_AddItem()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Search_cbo_Sort_AddItem"
    SubSkip = True

    With Search_cbo_Sort
        .Clear
        .AddItem "ID"
        .AddItem "Name"
        .AddItem "Department"
        .AddItem "Employment Status"
        .AddItem "Current Eligibility"
        .AddItem "INV Status"
        .Value = "Name"
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

Private Sub btn_AddNew_Click()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "btn_AddNew_Click"
    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module

    'TextBox_RowId.Value = ""
    UF_Employee.Show

ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub btn_Ascending_Click()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "btn_Ascending_Click"

    List_Sort xlAscending

ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub btn_Decsending_Click()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "btn_Decsending_Click"

    List_Sort xlDescending

ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub List_Sort(order As XlSortOrder)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "List_Sort"
    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    
    Dim ws1 As Worksheet, SHT1 As String
    Dim cb2v As String
    
    SHT1 = "Alerts"
    Set ws1 = ThisWorkbook.Sheets(SHT1)

    Select Case Search_cbo_Sort.Value
        Case "Name"
            cb2v = "I15"
        Case "ID"
            cb2v = "D15"
        Case "Department"
            cb2v = "S15"
        Case "Employment Status"
            cb2v = "T15"
        Case "Current Eligibility"
            cb2v = "AC15"
        Case "INV Status"
            cb2v = "AA15"
    End Select

    Select Case order
        Case 1
            ws1.Range("D14", Range("AK14").End(xlDown)).Sort Key1:=Range(cb2v), order1:=xlAscending, header:=xlYes
        Case 2
            ws1.Range("D14", Range("AK14").End(xlDown)).Sort Key1:=Range(cb2v), order1:=xlDescending, header:=xlYes
    End Select
    
ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
End Sub

Private Sub btn_Clear_Click()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "btn_Clear_Click"
    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    
    clear_sh 'Clear Sheet
    init_cbo 'Load ComboBoxes
    List_box_Data 'Load Data
    
ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    SubSkip = False
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub btn_Export_List_Click()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "btn_Export_List_Click"
    'SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    
    SubSkip = True

    Dim ws1 As Worksheet
    Dim ws2 As Worksheet, n As Long, SHT1 As String, SHT2 As String, RNG1 As Range, TBL1 As String, TBL2 As String, v As Long, w As Long, DbExtract As Variant
    
    SHT1 = "Alerts"
    SHT2 = "Roster"

    Set ws1 = ThisWorkbook.Sheets(SHT1)
    Set ws2 = ThisWorkbook.Sheets(SHT2)
    ws2.Cells.Clear
    n = ws1.Range("D" & Application.Rows.Count).End(xlUp).Row
    TBL1 = "D13:AK" & n
    TBL2 = "A1"
    
    Set RNG1 = ws1.Range(TBL1)
    
'    With ws
        RNG1.SpecialCells(xlCellTypeVisible).Copy
'        With .Range(TBL1).CurrentRegion.Select
'            Data = .Cells.Value
'        End With
'    End With
    ws2.Cells(1, 1).PasteSpecial
    ws2.Columns.AutoFit
'    Orkin "Sub Not Available"
'    .Value = "Return"
'    New_Edit.Show
    
ExitSub:
    'SwitchOff (True) 'On 'Disable VBA_Performance module
    SubSkip = False
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub btn_Home_Click()
    Page_CTRL -999
End Sub

Private Sub btn_Next_Click()
    Page_CTRL 25
End Sub

Private Sub btn_Prev_Click()
    Page_CTRL -25
End Sub

Private Sub Btn_Search_Click()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Btn_Search_Click"
    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    
    SubSkip = True
    List_box_Data 'Load Data
    
ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub checkbox_AutoFit_Change()
    Btn_Search_Click 'Update Data
End Sub

Private Sub checkbox_ShowCol_Change()
StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "checkbox_ShowCol_Change"

    If checkbox_ShowCol = True Then _
        checkbox_AutoFit.Value = True
        
    Btn_Search_Click 'Update Data

ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Search_cbo_Filter_Change()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Search_cbo_Filter_Change"
    If SubSkip = True Then
        Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
'    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    If Search_cbo_Filter = "" Then Search_cbo_Filter = "None"
    Search_cbo_Status_AddItem
    Btn_Search_Click 'Update Data
'    List_box_Data 'Load Data
    
ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Search_cbo_Sort_Change()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Search_cbo_Sort_Change"
    If SubSkip = True Then
        'Debug.Print "Skipping " & SubName & " from " & CallingSubName
        GoTo ExitSub
    End If
    If Search_cbo_Sort = "" Then Search_cbo_Sort = "Name"
    Btn_Search_Click 'Load Data
        
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Search_cbo_Status_Change()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Search_cbo_Status_Change"
    If SubSkip = True Then
        'Debug.Print "Skipping " & SubName & " from " & CallingSubName
        GoTo ExitSub
    End If
    If Search_cbo_Status = "" Then Search_cbo_Status = "All"
    Btn_Search_Click 'Load Data
        
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub Search_txt_Entry_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Search_txt_Entry_KeyDown"

    If KeyCode = 13 Then 'Keycode for "Enter" is 13
        'Search_txt_Entry_LostFocus
        Btn_Search_Click 'Load Data
    End If
    
ExitSub:
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

'Private Sub Search_txt_Entry_LostFocus()
'
'StartSub:
'    On Error GoTo ErrCtrl
'    CallingSubName = SubName
'    SubName = "Search_txt_Entry_LostFocus"
'    If SubSkip = True Then
'        'Debug.Print "Skipping " & SubName
'        GoTo ExitSub
'    End If
''    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
'
''    Search_cbo_Filter.Value = "None"
''    Orkin "Sub Not Available"
'
''    List_box_Data 'Load Data
''    Btn_Search_Click 'Load Data
'
'ExitSub:
''    SwitchOff (True) 'On 'Disable VBA_Performance module
'    SubName = CallingSubName
'    Exit Sub
'
'ErrCtrl:
'    ErrPrint Err.Number, Err.Description, SubName
'    Err.Clear
'    GoTo ExitSub
'
'End Sub

Private Sub Worksheet_Activate()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Worksheet(Alerts)_Activate"
    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    
    If SubSkip = True Then
        Debug.Print "Skipping " & SubName
        GoTo ExitSub
    End If

'Initialize worksheet and headers
    Orkin "Initialize worksheet and headers"
    SubSkip = True
    clear_sh 'Clear Sheet
    init_cbo 'Load ComboBoxes
    List_box_Data 'Load Data


ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubSkip = False
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Sub List_box_Data()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "List_box_Data"
'    SwitchOff (True) 'On 'VBA_Performance module
    
    ' Declare worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Alerts")
    
    ' Clear worksheet
'    With ws
        '.Protect "7704", UserInterfaceOnly:=True
'        If Not (qryRange Is Nothing) Then qryRange.Cells.Clear
'    End With
    
' Setup query
    Dim rst As New ADODB.Recordset
    Dim qrySELECT As String
    Dim qryFROM As String
    Dim qry As String, i As Integer
    Dim qryLIKE As String, n As Long
    Dim qrySORT As String
    Dim qryWHERE As String
    Dim qryWHAT As String
    Dim qryDIRECTION As String
    Dim qryNAME As String
    Dim qryFILTER As String
    
    qryDIRECTION = " ASC"
    qrySELECT = "*"
    
    If Search_cbo_Filter = "Alerts" And Search_cbo_Status = "All" Then
        qryFROM = "CLAlerts"
        qryWHERE = vbNullString
    ElseIf Search_cbo_Filter = "Alerts" Then
        qryFROM = "CLAlerts"
        qryWHERE = Concatenate(" WHERE ", NameAlias(Search_cbo_Status.Value))
    ElseIf Search_cbo_Filter = "None" Then
        qryFROM = "EmpDatabase"
        qryWHERE = vbNullString
    ElseIf Search_cbo_Status = "All" Then
        qryFROM = "EMPcurrent"
    Else
        qryFROM = "EmpDatabase"
        qryWHERE = Concatenate(" WHERE ", NameAlias(Search_cbo_Filter.Value, Search_cbo_Status.Value))
    End If
    
    'Double apostrophy if present in search box
    Search_txt_Entry = Replace(Search_txt_Entry.Value, "'", "''")
    
    If IsNullOrEmpty(Search_txt_Entry.Value) Then
        qryLIKE = vbNullString
    Else
        qryNAME = "AND File_As"
        If IsNullOrEmpty(qryWHERE) Then
            qryWHERE = " WHERE "
            qryNAME = "File_As"
        End If
        qryLIKE = Concatenate(qryNAME, " LIKE '%", Search_txt_Entry, "%'")
    End If
    
    qryNAME = NameAlias(Search_cbo_Sort.Value)
    
    If IsNull(qryNAME) Then GoTo qryConcat
    
    qrySORT = Concatenate(" ORDER BY  ", qryNAME, " ", qryDIRECTION)

qryConcat:
    qry = Concatenate("SELECT ", qrySELECT, " FROM ", qryFROM, qryWHERE, qryLIKE, qrySORT)
    qry = Replace(qry, "  ", " ")

    
ConnectDB:
    Orkin qry
    
' Connect to database
    rst.Open qry, ConnectDatabase, adOpenKeyset, adLockOptimistic
    
' Get # of Rows
    Dim rstCount As Long
    rstCount = CLng(rst.RecordCount)
    Dashboard.Range("Dashboard_RowCnt").Value = rstCount
    
' Get # of Columns
    Dim fldCount As Long
    fldCount = CLng(rst.Fields.Count)
    Dashboard.Range("Dashboard_ClmCnt").Value = fldCount
    
' Get Recordset data
    Dim rstDATA As Variant
    rstDATA = rst.GetRows
        
' Close connection
    rst.Close

' Format sheet & load headers
    TableHeaders
    
' Skip following if no results
    If rstCount = 0 Then
        ConditionalFormatting
        GoTo ExitSub
    End If

' Copy recordset to worksheet
    ws.Activate
    Dim rng As Range
    Application.Run "NamedRanges_Resize", "Alerts_qryRange", rstCount, fldCount
    Set rng = Range("Alerts_qryRange")

' Seperate Headers for listbox
    rstDATA = TransposeArray(rstDATA)
    PrintArray rstDATA, rng
    
CondFormat:
    ConditionalFormatting
    
LimitRows:
    Row_Limits
    GoTo ExitSub
    
ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    SubSkip = False
    Exit Sub

ErrCtrl:
    If Not Err.Number = 3021 Then
        ErrPrint Err.Number, Err.Description, SubName
        Err.Clear
        GoTo ExitSub
    Else
        Err.Clear
        Resume Next
    End If
    
End Sub
Private Sub Row_Limits()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Row_Limits"
'    SwitchOff (True) 'On 'VBA_Performance module

    Dim mainRng As Range
    Dim showRng As Range
    Dim rCount As Long
    Dim bRow As Long
    Dim tRow As Long
    Dim strRefers As String
    
    
    'Redefine
    tRow = Dashboard.Range("Dashboard_TopRow").Value
    bRow = Dashboard.Range("Dashboard_BtmRow").Value
    rCount = Dashboard.Range("Dashboard_RowCnt").Value
    strRefers = "=Alerts!$D$" & tRow & ":$AM$" & bRow
    ActiveWorkbook.Names.Item("Alerts_shwRange").RefersTo = strRefers
    
    Set mainRng = Alerts.Range("Alerts_qryRange")
    Set showRng = Alerts.Range("Alerts_shwRange")
    
    mainRng.EntireRow.Hidden = True
    showRng.EntireRow.Hidden = False

ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
        ErrPrint Err.Number, Err.Description, SubName
        Err.Clear
        GoTo ExitSub
    
End Sub
Sub Page_CTRL(inc As Long)
    Dim tRow As Long 'Top Row
    Dim bRow As Long 'Bottom Row
    Dim rCnt As Long 'Record Count

    
    btn_Prev.Enabled = True
    btn_Home.Enabled = True
    btn_Next.Enabled = True
    
    With Dashboard
        tRow = .Range("Dashboard_TopRow").Value + inc
        bRow = .Range("Dashboard_BtmRow").Value + inc
        rCnt = .Range("Dashboard_RowCnt").Value
        If Not tRow > 14 Then
            btn_Home.Enabled = False
            btn_Prev.Enabled = False
            tRow = 14
            bRow = 38
            Debug.Print "Not tRow > 14"
        ElseIf bRow > rCnt Then
            btn_Next.Enabled = False
            Debug.Print "bRow > rCnt"
        End If
        .Range("Dashboard_TopRow").Value = tRow
        .Range("Dashboard_BtmRow").Value = bRow
    End With

    Row_Limits
    
End Sub
'Sub Next_Click()
'
'    Dim tRow As Long 'Top Row
'    Dim bRow As Long 'Bottom Row
'    Dim rCnt As Long 'Record Count
'    Dim inc As Long 'Increment
'    inc = 25
'
'    With Dashboard
'        tRow = .Range("Dashboard_TopRow").Value + inc
'        bRow = .Range("Dashboard_BtmRow").Value + inc
'        rCnt = .Range("Dashboard_RowCnt").Value
'        .Range("Dashboard_TopRow").Value = tRow
'        .Range("Dashboard_BtmRow").Value = bRow
'
'    End With
'
'    btn_Prev.Enabled = True
'    btn_Home.Enabled = True
'
'    If bRow > rCnt Then btn_Next.Enabled = False
'
'    Row_Limits
'
'End Sub
'Sub Home_Click()
'
'    Dim tRow As Long
'    Dim bRow As Long
'
'    With Dashboard
'        tRow = 14
'        bRow = 38
'        .Range("Dashboard_TopRow").Value = tRow
'        .Range("Dashboard_BtmRow").Value = bRow
'    End With
'
'    btn_Prev.Enabled = False
'    btn_Home.Enabled = False
'    btn_Next.Enabled = True
'
'    Row_Limits
'
'End Sub
'Sub Prev_Click()
'
'    Dim tRow As Long 'Top Row
'    Dim bRow As Long 'Bottom Row
'    Dim rCnt As Long 'Record Count
'    Dim inc As Long 'Increment
'    inc = 25
'
'    With Dashboard
'        tRow = .Range("Dashboard_TopRow").Value - inc
'        bRow = .Range("Dashboard_BtmRow").Value - inc
'        rCnt = .Range("Dashboard_RowCnt").Value
'        If tRow < 14 Then
'            tRow = 14
'            btn_Prev.Enabled = False
'            btn_Home.Enabled = False
'        End If
'        .Range("Dashboard_TopRow").Value = tRow
'        .Range("Dashboard_BtmRow").Value = bRow
'    End With
'
'    btn_Home.Enabled = True
'
'    Row_Limits
'
'End Sub

Private Sub clear_sh()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "clear_sh"
'    SwitchOff (True) 'On 'VBA_Performance module
        
    Dim n As Integer
    Dim ole As OLEObject
    
    btn_Home_Click
    
    With Worksheets("Alerts")
'        .Protect "7704", UserInterfaceOnly:=True
        If Not (Range("Alerts_qryRange") Is Nothing) Then Range("Alerts_qryRange").Clear
        .Search_txt_Entry.Value = vbNullString
        .Search_lbl_RCount.Caption = vbNullString
        For Each ole In .OLEObjects
            ole.Visible = False
        Next
    End With
        
ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub
Private Sub TableHeaders()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "TableHeaders"
'    SwitchOff (True) 'On 'VBA_Performance module
        
' Declare Worksheet
    Dim ws As Worksheet
    Set ws = Alerts
    
    Dim rowCount As Long
    Dim clmCount As Long
    
    rowCount = Dashboard.Range("Dashboard_RowCnt").Value
    clmCount = Dashboard.Range("Dashboard_ClmCnt").Value

'Worksheet Format
    With ws
    'Worksheet Header
    Application.Run "NamedRanges_Resize", "Alerts_sheetHeader", 11, clmCount
'        Dim WrkShtHdrRng As Variant
'        Set WrkShtHdrRng = ws.Range("D3")
'        Dim WrkShtHdrSize As Range
'
'        WrkShtHdrRng.CurrentRegion = Range("B3").Value
        'WrkShtHdrRng.Select

        Dim WrkShtHdr As Range
        Set WrkShtHdr = Range("Alerts_sheetHeader")
        With WrkShtHdr
            With .Borders
                .LineStyle = xlContinuous
                .ColorIndex = 2
                .TintAndShade = 0
                .Weight = xlThin
            End With
            .Merge
            .Interior.Color = rgb(75, 0, 75)
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlTop
            With .Font
                .Name = "Arial Rounded MT Bold"
                .Size = "30"
                .Bold = True
                .Color = vbWhite
            End With
        End With
        
    'Worksheet Lables and Controls
        Dim ole As OLEObject
        For Each ole In ws.OLEObjects
            SL_Format ole
            ole.Visible = True
            ole.Placement = 3
        Next
        checkbox_AutoFit.Visible = False
        lbl_Welcome.Caption = "Welcome, " & Dashboard.Range("Dashboard_Username").Value

    'Table Header
        Dim tblHead As Range
        Application.Run "NamedRanges_Resize", "Alerts_qryHeaders", 1, clmCount
        Set tblHead = Range("Alerts_qryHeaders")
        Dim headers As Variant
        Dim ColumnWidth As Variant
        headers = Array("ID", "First", "Middle", "Last", "Suffix", "Full Name", "SSN", "DOB", "City", "State", "Address", "Cellular", "Email", "Hire Date", "Term Date", "Department", "Employment Status", "FP Submit", "SAC", "DISS", "INV Open", "INV Close", "INV Type", "INV Status", "INV Due", "Eligibility", "Eligibility Date", "CE", "NDA", "Access", "ALTESS", "Remove", "Type", "Debrief", "Site", "Tag")
        
    'Show all Columns
        If checkbox_ShowCol = True Then
            checkbox_AutoFit.Visible = True
            ColumnWidth = Array(15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15)
        Else
    'Show only working Columns
            ColumnWidth = Array(5, 0, 0, 0, 0, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 15, 20, 0, 0, 0, 15, 0, 0, 20, 15, 15, 15, 15, 15, 15, 0, 0, 0, 0, 0, 0)
            checkbox_AutoFit.Value = False
            checkbox_AutoFit.Visible = False
        End If
        
        With tblHead
            .HorizontalAlignment = xlCenter
            .Interior.Color = rgb(192, 192, 192)
            .Value = headers
            .ColumnWidth = ColumnWidth
            With .Font
                .Size = "12"
                .Bold = True
                .Color = vbBlack
            End With
            With .Borders
                .LineStyle = xlContinuous
                .ColorIndex = 2
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
    End With
    
ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Worksheet_BeforeDelete()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Worksheet(""Alerts"")_BeforeDelete"
'    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module

Dim MyName As String

    MyName = ThisWorkbook.ActiveSheet.Name 'Capture the original worksheet name
    With ThisWorkbook.ActiveSheet
        .Name = Left(MyName, 30) + "#" 'Rename the worksheet
        .Copy After:=Sheets(ThisWorkbook.ActiveSheet.Index) 'Create a copy of the worksheet
        .Name = MyName 'Name the copy to the original name
    End With
    
ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub Worksheet_Deactivate()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Worksheet(""Alerts"")_Deactivate"
'    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module

'    Call clear_sh


ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Worksheet(""Alerts"")_SelectionChange"
'    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module

    If SubSkip = True Then
        'Orkin "Skipping from " & CallingSubName
        GoTo ExitSub
    End If
    
    Dim ws As Worksheet
    Dim tr As Integer
    Dim n As Integer
    'Dim rowCount As Integer
    
    Dim objColorStop As ColorStop
    
    Set ws = Alerts
    On Error GoTo ExitSub
    'rowCount = ws.Range("D" & Application.Rows.Count).End(xlUp)
    On Error GoTo ErrCtrl
    
    n = ws.Range("D" & Application.Rows.Count).End(xlUp).Row
    tr = Target.Row
    
    'If Not Intersect(Target, ws.Range("D" & Application.Rows.Count).End(xlUp)) Is Nothing And Range("D" & Target.Row).Value <> Empty Then
    'If Range("D" & Target.Row).Value <> Empty Then
    'If Not Intersect(Target, qryRange) Is Nothing Then
    'If Not Target.Row > 13 Then
    If Not Intersect(Target, ws.Range("D" & Target.Row, "AM" & n)) Is Nothing Then
        SubSkip = True
        Dashboard.Range("Dashboard_TargetRow").Value = Target.Row
        'tr = Range("B5").Value
    Else
        Range("D2").Select
        GoTo ExitSub
    End If
    
    If tr < 14 Then
        SubSkip = True
        Range("D" & 14 & ":AM" & 14).Select
'        n = 14
'        With Range("D" & n & ":AL" & n)
'            .Interior.Gradient.Degree = 90
'            .Interior.Gradient.ColorStops.Clear
'            .Interior.Pattern = xlPatternLinearGradient
'            Set objColorStop = .Interior.Gradient.ColorStops.Add(0)
'            objColorStop.Color = vbYellow
'            Set objColorStop = Range("A1").Interior.Gradient.ColorStops.Add(0.5)
'            objColorStop.Color = vbGreen
'        End With
        
        'TextBox_Info.Value = "1"
    'ElseIf Target.Address = Range("D" & tr & ":AM" & tr) Then
    'ElseIf tr >= 14 And tr <= n Then

        
'        With Range("D" & n & ":AL" & n)
'            .Interior.Gradient.Degree = 90
'            .Interior.Gradient.ColorStops.Clear
'            .Interior.Pattern = xlPatternLinearGradient
'            Set objColorStop = .Interior.Gradient.ColorStops.Add(0)
'            objColorStop.Color = vbYellow
'            Set objColorStop = Range("A1").Interior.Gradient.ColorStops.Add(0.5)
'            objColorStop.Color = vbGreen
'        End With
    ElseIf tr > n Then
        SubSkip = True
        Range("D" & n & ":AM" & n).Select
        'TextBox_Info.Value = "3 (" & Range("A" & n) & ") (" & Range("E" & n) & ") (" & Range("N" & n) & ")"
        Dashboard.Range("Dashboard_TargetID").Value = Range("D" & n & ":AM" & n).Row
    Else
        SubSkip = True
        Range("D" & tr & ":AM" & tr).Select
        Dashboard.Range("Dashboard_TargetID").Value = Range("D" & tr)
        btn_Export_List.Enabled = True
    End If
    ''Debug.print "(" & Target.Row&; ", " & n & "): " & TextBox_Info.Value


ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubSkip = False
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub ConditionalFormatting()

StartSub:
On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "ConditionalFormatting"
    'SwitchOff (True) 'On 'Enable VBA_Performance module
    
    Dim entry As Range
    Dim rCount As Long
    
    rCount = Dashboard.Range("Dashboard_RowCnt").Value
    If Not rCount = 1 Then
        Search_lbl_RCount.Caption = Format(rCount, "#,##0") & " Records Found"
    Else
        Search_lbl_RCount.Caption = Format(rCount, "#,##0") & " Record Found"
    'On Error Resume Next
    End If
    
    Application.StatusBar = Search_lbl_RCount.Caption
    
    If rCount = 0 Then
        SubSkip = True
'        Range("D2").Select
        GoTo ExitSub
    End If
    
' CREATE ALL THE CONDITIONAL FORMATTING RULES:
' Format query range
    Dim tblQuery As Range
    Set tblQuery = Alerts.Range("Alerts_qryRange")
        With tblQuery
            .HorizontalAlignment = xlCenter
            .Interior.Color = rgb(221, 235, 247) 'Sets to lt. blue
            With .Borders
                .LineStyle = xlContinuous
                .ColorIndex = 2
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End With
' Format each entry
    For Each entry In tblQuery
        With entry
    ' Change Full Name field back
            Range("I" & .Row).HorizontalAlignment = xlLeft
        ' For active employees only
            If Range("T" & .Row) = "Terminated" Then
                .Interior.Color = rgb(242, 242, 242) 'Sets to lt. gray
                .Font.ColorIndex = 48
                .Font.Italic = True
                .Font.Bold = False
        ' PSQ Over Due
            ElseIf Range("AA" & .Row) <> "Sent to ISP" And Range("AA" & .Row) <> "FSO Review" _
            And Range("AA" & .Row) <> "Needs NDA" And CDate(Range("AB" & .Row)) < Date _
            And Range("AB" & .Row) <> "" Then
            ' Always highlight paper version
                If Range("AA" & .Row) = "Paper Version" Then
                    .Interior.Color = rgb(255, 0, 0)
                    .Font.Bold = False
                    .Font.Color = vbBlack
                Else
            ' Over Due
                    .Interior.Color = rgb(255, 255, 0) 'RGB(192, 0, 0)
                    .Font.Bold = False
                    .Font.Color = vbBlack
                End If
                
        ' Needs to sign NDA
            ElseIf Range("AA" & .Row) = "Needs NDA" Then
                .Interior.Color = rgb(0, 0, 255)
                .Font.Bold = False
                .Font.Color = vbWhite
        
        ' PSQ Terminated
            ElseIf Range("AA" & .Row) = "PSQ Terminated" Then
                .Interior.Color = rgb(0, 0, 0)
                .Font.Bold = False
                .Font.Color = vbRed
        
        ' FSO Review
            ElseIf Range("AA" & .Row) = "FSO Review" Then
                .Interior.Color = rgb(255, 0, 0)
                .Font.Bold = False
                .Font.Color = vbBlack
        
        ' Need Prints
            ElseIf Range("AC" & .Row) = "None" And Range("U" & .Row) = "" Then
                Range("AM" & .Row) = "Need Prints"
                .Interior.Color = rgb(112, 48, 160)
                .Font.Bold = False
                .Font.Color = vbWhite
        
        ' PR Due
            ElseIf Range("AD" & .Row) <> "" And Range("Z" & .Row) <> "Tier 3R" And Range("AD" & .Row) _
            < Date - 1642 And Range("AE" & .Row) < Date - 1642 Then
                Range("AM" & .Row) = "PR Due"
                .Interior.Color = rgb(192, 0, 0)
                .Font.Bold = True
                .Font.Color = vbYellow
              
        ' Eligibility Pending
            ElseIf Range("AC" & .Row) = "Elig Pending" Then
                .Interior.Color = rgb(255, 192, 0)
                .Font.Bold = False
                .Font.Color = vbBlack
                
        ' Needs Access in Diss
            ElseIf Range("T" & .Row) = "Current" And Range("AF" & .Row) <> "" And Range("AG" & .Row) = "" Then
                Range("AM" & .Row) = "Needs Access"
                .Interior.Color = rgb(169, 208, 142)
                .Font.Bold = False
                .Font.Color = vbBlack
            End If
        'Autofit columns in range
            If checkbox_AutoFit Then Columns.AutoFit
        End With
    Next entry
    
ExitSub:
    'SwitchOff (True) 'On 'Disable VBA_Performance module
    'On Error GoTo 0
    SubSkip = False
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Private Sub init_cbo() 'Initialize ComboBoxes

StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "init_cbo"
'   SwitchOff (True) 'On 'Enable VBA_Performance module

    Search_cbo_Filter_AddItem
    Search_cbo_Status_AddItem
    Search_cbo_Sort_AddItem

ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub SL_Format(slName As Object) 'Format search labels

StartSub:
   On Error Resume Next
   CallingSubName = SubName
   SubName = "SL_Format"
'   SwitchOff (True) 'On 'Enable VBA_Performance module

    If InStr(slName.progID, "CommandButton") Then
        GoTo FormatButton
    ElseIf Not Left(slName.Name, 7) = "Search_" Then
        Orkin "Format(A): " & slName.Name
        GoTo FormatA
    Else
        Orkin "Format(B): " & slName.Name
        GoTo FormatB
    End If

    
FormatA: 'Purple Background
    With slName
        .Height = "18"
        With .Object
            .BackColor = &H4B004B
            .BackStyle = 0
            .ForeColor = &H80000005
            .TextAlign = 2
            .SpecialEffect = 0
            If InStr(slName.progID, "CheckBox") Then GoTo FormatCheckBox
            .BorderColor = &H80000006
            .BorderStyle = 0
        End With
    End With
    GoTo FormatFont
    
FormatB:
    With slName.Object
        .BackColor = &H80000005
        .ForeColor = &H80000008
        .BackStyle = 1
        .TextAlign = 2
        .SpecialEffect = 0
        .BorderColor = &H80000006
        .BorderStyle = 0
        If slName.Name = "Search_lbl_RCount" Then .Caption = Format(0, "#,##0") & " Records"
        If slName.Name = "Search_txt_Entry" Then .TextAlign = 1
    End With
    GoTo FormatFont
    
FormatButton:
    With slName
        .Height = 24
        .Width = 72
        With btn_Ascending
                .Height = 18
                .Width = 18
        End With
        With btn_Decsending
                .Height = 18
                .Width = 18
        End With
        With .Object
            .BackColor = &HFFFFFF
            .ForeColor = &H80000001
            .AutoSize = False
            With .Font
                .Name = "Calibri"
                .Size = "12"
                .Bold = True
            End With
        End With
    End With
    GoTo ExitSub
    
FormatFont:
    With slName.Object.Font
        .Name = "Calibri"
        .Size = "12"
        .Bold = True
    End With
    GoTo ExitSub
        
FormatCheckBox:
    With slName
        With .Border
            .Color = &H4B004B
            .LineStyle = 0
        End With
        '.Left = "899.25"
        '.Top = "165.75"
        With .Object.Font
            .Name = "Calibri"
            .Size = "8"
            .Bold = True
            '.Caption = "Show All Columns"
        End With
    End With
    GoTo ExitSub
    
ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, slName.Name
    Err.Clear
    GoTo ExitSub
    
End Sub
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Worksheet(""Alerts"")_BeforeDoubleClick"
'    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    
    Dim tagID As Integer
    Dim trange As Range
    
    Cancel = True 'This is to stop the user from entering in-cell edit mode
    
    With ThisWorkbook.Worksheets("Alerts")
        Set trange = .Range("D15:AK15", .Range("D" & .Rows.Count).End(xlUp))
    End With
    
'    If Not Target.Row >= 15 Then GoTo ExitSub 'outside selectable range
    
    If Not Intersect(Target, trange) Is Nothing Then
        tagID = Cells(ActiveCell.Row, 4)
        With UF_Employee
            .Tag = tagID '<~~ tell the UserForm there's something to bring in so that it'll fill controls from the sheet instead of initializing them
            .Show
            '.Tag = "" '<~~ bring Tag property back to its "null" value
        End With
    Else
        GoTo ExitSub 'outside selectable range
    End If
    
    'Refresh list
'    List_box_Data 'Load Data
    Btn_Search_Click 'Load Data
    

ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub
Sub ShowCol(shCol As Boolean)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "ShowCol"
'    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    Dim shAC As String
    
    ShwAllColumns = shCol
    
    If shCol = True Then
        shAC = "Showing all columns."
    Else
        shAC = "Showing limited columns."
    End If
    
    Btn_Search_Click
    
    Debug.Print "******" & vbCrLf & shAC

ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub

End Sub

Sub PrintArray(Data As Variant, cl As Range)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "PrintArray"
    
'    SwitchOff (True) 'On 'VBA_Performance module 'Enable VBA_Performance module
    'Cl.Resize(UBound(Data), 1).Value = Application.Transpose(Data)
    cl.Resize(UBound(Data, 1) + 1, UBound(Data, 2)) = Data
    'Debug.Print Data
    
ExitSub:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Function NameAlias(Filter As String, Optional Status As String) As String

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "NameAlias"
    
    Dim str As String
    Dim stat As Boolean
    If Not IsNullOrEmpty(Status) Then stat = True
    
    Select Case Filter
        Case ""
            str = vbNullString
        Case "ID"
            str = "db_ID"
        Case "First"
        Case "Middle"
        Case "Last"
        Case "Suffix"
        Case "Full Name", "Name"
            str = "File_As"
'        Case "Name"
'            str = "File_As"
        Case "SSN"
        Case "DOB"
        Case "City"
        Case "State"
        Case "Address"
        Case "Cellular"
        Case "Email"
        Case "Hire Date"
        Case "Term Date"
        Case "Department"
            str = "cbo_DEPT"
        Case "Employment Status"
            str = "cbo_EMPstatus"
        Case "FP Submit"
        Case "SAC"
        Case "DISS"
        Case "INV Open"
        Case "INV Close"
        Case "INV Type"
        Case "INV Status"
            str = "cbo_PSQstatus"
        Case "INV Due"
        Case "Current Eligibility", "Eligibility"
            str = "cbo_ELIGstatus"
'        Case "Eligibility"
'            str = "cbo_ELIGstatus"
        Case "Eligibility Date"
        Case "CE"
        Case "NDA"
        Case "Access"
        Case "ALTESS"
        Case "Remove"
        Case "Type"
        Case "Debrief"
        Case "Site"
        Case "Tag"
        Case "PSQ Over Due"
        Case "Fingerprints Needed"
            str = "cbo_PSQstatus"
            stat = True
        Case "Needs Review"
        Case "Ready for Release"
        Case "Needs NDA"
        Case "PSQ Terminated"
            str = "cbo_PSQstatus"
            stat = True
        Case "Eligibility Pending"
        Case "PR Due"
            str = "cbo_PSQstatus"
            stat = True
    End Select

    If Not stat Then
        NameAlias = str
    Else
        If IsNullOrEmpty(Status) Then Status = Filter
        NameAlias = Concatenate(str, " = ", "'", Status, "'")
    End If
    
    Orkin NameAlias
    
ExitSub:
    SubName = CallingSubName
    Exit Function

ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
End Function
