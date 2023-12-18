Attribute VB_Name = "SheetPrep"
Option Explicit
Dim show As Boolean

Public Sub Sheet_Prep()
#If Dev = True Then
    show = True
#Else
    show = False
#End If

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "Sheet_Prep"
    'SwitchOff (True) 'VBA_Performance module
    
    SubSkip = True 'Skip Sheet Activate

    Dim imageFolder As String
    Dim imageFile As String
    Dim bPic As String
    
    imageFolder = Application.ActiveWorkbook.Path & Application.PathSeparator & "Images"
    imageFile = "purple_neon_abstract_4k.jpg"
    bPic = imageFolder & Application.PathSeparator & imageFile
    
' Delete named ranges
    NamedRanges_Delete
       
' Declare Worksheet
    Dim ws As Worksheet
' Loop through each sheet
    For Each ws In ActiveWorkbook.Worksheets
    ' Hide page breaks
        ActiveWindow.DisplayGridlines = False
    ' With worksheet
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
        ' Set developer range
            Dim DevRange As Range
            Set DevRange = Range("A1:B40")
            With DevRange
                .EntireColumn.Hidden = True
                .Name = ws.Name & "_" & "DevRange"
                .ColumnWidth = Array(10, 22)
                .Interior.Color = rgb(58, 56, 56)
                With .Font
                    .Size = "8"
                    .Bold = False
                    .Color = vbWhite
                End With
            End With
            
        'Developer values
            Dim r As Range
            Dim i As Integer
            Dim clmA As Variant
            Dim clmB As Variant
            Dim rAddress As String
            Dim shRef As String
            i = 0
            clmA = Array("Username", "Title", "Page", "Row Cnt", "Clm Cnt", "Target Row", "Target ID", "Top Row", "Btm Row")
            clmB = Array(GetUserName, "EMPLOYEE CLEARANCE TRACKER", .CodeName, vbNullString, vbNullString, vbNullString, vbNullString, 14, 38)
            shRef = "=Dashboard!"
            For Each r In Range("A2:B" & (Application.CountA(clmA) + 1))
                rAddress = Replace(r.Address, "$", vbNullString)
                If Left(rAddress, 1) = "A" Then r.Value = clmA(i)
                If Left(rAddress, 1) = "B" And ws.CodeName = "Dashboard" Then
                    r.Value = clmB(i)
                    r.Name = ws.Name & "_" & Replace(Range(Replace(rAddress, "B", "A")).Value, " ", vbNullString)
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
                ElseIf ws.CodeName = "Roster" Then
                    With ws
                        .Range("C2").Name = ws.Name & "_" & "rptHeaders"
                        .Range("C3").Name = ws.Name & "_" & "rptRange"
                    End With
                End If
        'Set page header
            Dim WrkShtHdrSize As Range
            Set WrkShtHdrSize = Range("D2:Z11")
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
Call NamedRanges_Show
Dashboard.Activate

    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
End Sub

Sub NamedRanges_Show()

StartSub:
        On Error GoTo ErrCtrl
        CallingSubName = SubName
        SubName = "ShowNamedRanges"
        SwitchOff (True) 'VBA_Performance module
    
        Dim MyName As Name
        
        For Each MyName In Names
        
        ActiveWorkbook.Names(MyName.Name).Visible = show
        
        Next
ExitSub:
        SwitchOff (False) 'Disable VBA_Performance module
        SubName = CallingSubName
        Exit Sub
        
ErrCtrl:
        ErrPrint Err.Number, Err.Description, SubName
        Err.Clear
        GoTo ExitSub
        
    End Sub

Sub NamedRanges_Delete()

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "DeleteNamedRanges"
    SwitchOff (True) 'VBA_Performance module

    Dim MyName As Name
    
    For Each MyName In Names
    
    ActiveWorkbook.Names(MyName.Name).Delete
    
    Next
ExitSub:
    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitSub
    
End Sub

Sub NamedRanges_Resize(strNamedRange As String, row As Long, Column As Long)

StartSub:
    On Error GoTo ErrCtrl
    CallingSubName = SubName
    SubName = "NamedRanges_Resize"
    SwitchOff (True) 'VBA_Performance module
    
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    
    Dim NamedRange As Name
    'Set NamedRange = wb.Names.Item(strNamedRange)
    
'    With NamedRange
'        .RefersTo = .RefersToRange.Resize(row, Column)
'    End With

ExitSub:
    SwitchOff (False) 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Sub
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    Resume Next
    'GoTo ExitSub
    
End Sub
