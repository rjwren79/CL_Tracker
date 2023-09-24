Attribute VB_Name = "SheetPrep_Mod"
Option Explicit

Public Sub SheetPrep(Optional ShowAdminRng As Boolean) 'Show Admin Range

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Sheet_Prep"
    SwitchOff (True) 'On 'VBA_Performance module
    
    SubSkip = True 'Skip Sheet Activate
    Dim sar As Boolean
    
    sar = Not ShowAdminRng

    With Application
    ' Hide scroll bars
        .DisplayScrollBars = sar
    ' Hide Formula bar
        .DisplayFormulaBar = sar
    End With
    
    ' Delete named ranges
        NamedRanges_Delete
    ' Hide page breaks
        ActiveWindow.DisplayGridlines = False

' Declare background image
    Dim imageFolder As String
    Dim imageFile As String
    Dim bPic As String
    imageFolder = Application.ActiveWorkbook.Path & Application.PathSeparator & "Images"
    imageFile = "purple_neon_abstract_4k.jpg"
    bPic = imageFolder & Application.PathSeparator & imageFile
       
' Declare Worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        With ws
            .Activate
        ' Disable page breaks
            .DisplayPageBreaks = False
        ' Set background image
            .SetBackgroundPicture FileName:=bPic
        ' Format cells
            With Cells
                .Clear
                .RowHeight = "15"
                .ColumnWidth = "8.43"
                With .Font
                    .Size = "12"
                    .Bold = False
                    .Color = vbBlack
                End With
            End With
        ' Set admin range
            Dim AdminRange As Range
            Set AdminRange = Range("A1:B40")
            With AdminRange
                .Name = ws.Name & "_" & "AdminRange"
                .ColumnWidth = Array(10, 22)
                .Interior.Color = rgb(58, 56, 56)
                .EntireColumn.Hidden = sar
                With .Font
                    .Size = "8"
                    .Bold = False
                    .Color = vbWhite
                End With
            End With
            
        ' Admin values
            Dim r As Range
            Dim i As Integer
            Dim clmA As Variant
            Dim clmB As Variant
            Dim rAddress As String
            Dim shRef As String
            i = 0
            clmA = Array("Username", "Title", "Page", "Row Cnt", "Clm Cnt", "Target Row", "Target ID", "Top Row", "Btm Row")
            clmB = Array(GetUserName, "EMPLOYEE CLEARANCE TRACKER", .CodeName, "", "", "", "", 14, 38)
            shRef = "=Dashboard!"
            For Each r In Range("A2:B" & (Application.CountA(clmA) + 1))
                rAddress = Replace(r.Address, "$", "")
                If Left(rAddress, 1) = "A" Then r.Value = clmA(i)
                If Left(rAddress, 1) = "B" And ws.CodeName = "Dashboard" Then
                    r.Value = clmB(i)
                    r.Name = ws.Name & "_" & Replace(Range(Replace(rAddress, "B", "A")).Value, " ", "")
                    i = i + 1
                ElseIf Left(rAddress, 1) = "B" Then
                    If rAddress = "B4" Then
                        r.Value = ws.CodeName
                    Else
                        r.Value = shRef & r.Address
                    End If
                    i = i + 1
                End If
            Next
            If ws.CodeName = "Alerts" Then
                With ws
                    .Range("D13").Name = ws.Name & "_" & "qryHeaders"
                    .Range("D14").Name = ws.Name & "_" & "qryRange"
                    .Range("D14").Name = ws.Name & "_" & "shwRange"
                    .Range("I14").Name = ws.Name & "_" & "fullName"
                End With
            End If
        ' Set page header
            Dim WrkShtHdrSize As Range
            Set WrkShtHdrSize = Range("D2:AA12")
            With WrkShtHdrSize
                .Name = ws.Name & "_" & "sheetHeader"
                With .Borders
                    .LineStyle = xlContinuous
                    .ColorIndex = 2
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With .Font
                    .Name = "Arial Rounded MT Bold"
                    .Size = "30"
                    .Bold = True
                    .Color = vbWhite
                End With
                .Merge
                .Interior.Color = rgb(75, 0, 75)
                .HorizontalAlignment = xlCenterAcrossSelection
                .VerticalAlignment = xlTop
                .Value = Range("Dashboard_Title").Value
            End With
        End With
    Next
    
ExitSub:
    'Dashboard.Activate
    Alerts.Activate
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
End Sub
Public Function SheetPrep_Help() As String

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "SheetPrep_Help"
    SwitchOff (True) 'On 'VBA_Performance module
    
    SheetPrep_Help = "The SheetPrep function was designed to format all sheets with in the application." & Chr(10) & "Parameters:" & Chr(10) _
    & "  Optional ShowAdminRng As Boolean   :   By default the admin columns are hidden. Add True agument to bypass default." & Chr(10) _
    & " Contact the R.Wren at richard.wren@brdnest.net for more information."
    
ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Function
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
End Function
Private Sub UIctrls()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "UIctrls"
    SwitchOff (True) 'On 'VBA_Performance module
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .WindowState = xlMaximized
        .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",False)"
        .CommandBars("Full Screen").Visible = False
        .CommandBars("Worksheet Menu Bar").Enabled = False
        .DisplayStatusBar = False
        .DisplayScrollBars = False
        .DisplayFormulaBar = False
'        .Width = 800
'        .Height = 450
    End With
    With ActiveWindow
        .DisplayWorkbookTabs = False
        .DisplayHeadings = False
        .DisplayRuler = False
        .DisplayFormulas = False
        .DisplayGridlines = False
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    
ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
End Sub
Private Sub NamedRanges_Delete()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "NamedRanges_Delete"
    SwitchOff (False) 'VBA_Performance module

    Dim MyName As Name
    
    For Each MyName In Names
    
    ActiveWorkbook.Names(MyName.Name).Delete
    
    Next
ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Private Sub NamedRanges_Resize(strNamedRange As String, Row As Long, Column As Long)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "NamedRanges_Resize"
    SwitchOff (True) 'On 'VBA_Performance module
    
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    
    Dim NamedRange As Name
    Set NamedRange = wb.Names.Item(strNamedRange)
    
    With NamedRange
        .RefersTo = .RefersToRange.Resize(Row, Column)
    End With

ExitSub:
    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub
