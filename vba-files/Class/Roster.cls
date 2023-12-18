VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Roster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub TableHeaders()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "TableHeaders"
'    SwitchOff (True) 'VBA_Performance module
        
' Declare Worksheet
    Dim ws As Worksheet
    Set ws = Roster
    
    Dim rowCount As Long
    Dim clmCount As Long
    
'    rowCount = Dashboard.Range("Dashboard_RowCnt").Value
'    clmCount = Dashboard.Range("Dashboard_ClmCnt").Value

'Worksheet Format
    With ws
    'Worksheet Header
    NamedRanges_Resize "Roster_sheetHeader", 11, clmCount
'        Dim WrkShtHdrRng As Variant
'        Set WrkShtHdrRng = ws.Range("D3")
'        Dim WrkShtHdrSize As Range
'
'        WrkShtHdrRng.CurrentRegion = Range("B3").Value
        'WrkShtHdrRng.Select

'        Dim WrkShtHdr As Range
'        Set WrkShtHdr = Range("Roster_sheetHeader")
'        With WrkShtHdr
'            With .Borders
'                .LineStyle = xlContinuous
'                .ColorIndex = 2
'                .TintAndShade = 0
'                .Weight = xlThin
'            End With
'            .Merge
'            .Interior.Color = rgb(75, 0, 75)
'            .HorizontalAlignment = xlCenterAcrossSelection
'            .VerticalAlignment = xlTop
'            With .Font
'                .Name = "Arial Rounded MT Bold"
'                .Size = "30"
'                .Bold = True
'                .Color = vbWhite
'            End With
'        End With
        
    'Worksheet Lables and Controls
'        Dim ole As OLEObject
'        For Each ole In ws.OLEObjects
'            SL_Format ole
'            ole.Visible = True
'            ole.Placement = 3
'        Next
'        checkbox_AutoFit.Visible = False
'        Dim usrnm As String
'        usrnm = Dashboard.Range("Dashboard_Username").Value
'        lbl_Welcome.Caption = "Welcome, " & usrnm
'        usrnm = vbNullString
    'Table Header
        Dim tblHead As Range
        NamedRanges_Resize "Roster_qryHeaders", 1, clmCount
        Set tblHead = Range("Roster_qryHeaders")
        Dim headers As Variant
        Dim ColumnWidth As Variant
        headers = Array("ID", "First", "Middle", "Last", "Suffix", "Full Name", "SSN", "DOB", "City", "State", "Address", "Cellular", "Email", "Hire Date", "Term Date", "Department", "Employment Status", "FP Submit", "SAC", "DISS", "INV Open", "INV Close", "INV Type", "INV Status", "INV Due", "Eligibility", "Eligibility Date", "CE", "NDA", "Access", "ALTESS", "Remove", "Type", "Debrief", "Site", "Tag")
        
    'Show all Columns
'        If checkbox_ShowCol = True Then
'            checkbox_AutoFit.Visible = True
'            ColumnWidth = Array(15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15)
'            'If checkbox_AutoFit Then Set ColumnWidth = Columns.AutoFit
'        Else
    'Show only working Columns
            ColumnWidth = Array(5, 0, 0, 0, 0, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 15, 20, 0, 0, 0, 15, 0, 0, 20, 15, 15, 15, 15, 15, 15, 0, 0, 0, 0, 0, 0)
            checkbox_AutoFit.Value = False
            checkbox_AutoFit.Visible = False
'        End If
        
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
'    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub Worksheet_Activate()
    'TableHeaders
End Sub
