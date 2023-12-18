Attribute VB_Name = "VBA_Performance"
Option Explicit
Dim lCalcSave As Long
Dim bScreenUpdate As Boolean
Public SheetIni As Boolean
Public ShowOrkin As Boolean
Public ShowOrkinErr As Boolean
Public SubSkip As Boolean
Public SubName As String
Public CallingSubName As String

Sub Show_Window(sw As Boolean)

    ThisWorkbook.Activate
    ActiveWindow.Visible = sw
    
End Sub


Sub terminix(shOrkin As Boolean)

StartSub:
    On Error GoTo ExitSub
    Dim shwO As String
    
    ShowOrkin = shOrkin
    
    If shOrkin = True Then
        shwO = "Orkin reporting."
    Else
        shwO = "Orkin not reporting."
    End If
    
    Debug.Print "******" & vbCrLf & shwO
    
ExitSub:
    Exit Sub

End Sub

Sub SwitchOff(bSwitchOff As Boolean)

StartSub:
    On Error GoTo ExitSub
    
    Dim ws As Worksheet
    
    With Application
        If bSwitchOff Then
            ' On
            Orkin "Performance Mode: ON"
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            lCalcSave = .Calculation
            bScreenUpdate = .ScreenUpdating
            
            
            .EnableAnimations = False
            '
        Else
            ' Off
            If .Calculation <> lCalcSave And lCalcSave <> 0 Then .Calculation = lCalcSave
                .ScreenUpdating = bScreenUpdate
                .EnableAnimations = True
                Orkin "Performance Mode: OFF"
        End If
    End With
    
ExitSub:
    Exit Sub

End Sub
Sub ForceScreenUpdate()

    Orkin "Force Screen Update"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
   ' Application.Wait Now + #12:00:01 AM#
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    SwitchOff (True)
    
End Sub
Sub ErrPrint(errNum As Integer, errDes As String, ProcName As String)

    Dim dbPrint As String
    ShowOrkinErr = True
    dbPrint = "Error in process: " & ProcName & vbCrLf & "Error #: " & Err.Number & vbCrLf & "Error Description: " & Err.Description
    Orkin dbPrint
    ShowOrkinErr = False
    Exit Sub
    
End Sub

Sub Orkin(txt As String)

    Dim sMsg As String
    sMsg = "************************" & vbCrLf & "Sub Name: " & SubName & vbCrLf & txt & vbCrLf

    If ShowOrkinErr = True Then
        ' Popup display the message
        MsgBox sMsg, Title:="Error"
    Else
        Debug.Print sMsg
        Exit Sub
    End If

End Sub

Sub ProgressMeter()

Dim booStatusBarState As Boolean
Dim fractionDone As Integer
Dim iMax As Integer
Dim i As Integer

iMax = 10000

    Application.ScreenUpdating = False
''//Turn off screen updating

    booStatusBarState = Application.DisplayStatusBar
''//Get the statusbar display setting

    Application.DisplayStatusBar = True
''//Make sure that the statusbar is visible

    For i = 1 To iMax ''// imax is usually 30 or so
        fractionDone = CDbl(i) / CDbl(iMax)
        Application.StatusBar = Format(fractionDone, "0%") & " done..."
        ''// or, alternatively:
        ''// statusRange.value = Format(fractionDone, "0%") & " done..."
        ''// Some code.......

        DoEvents
        ''//Yield Control

    Next i

    Application.DisplayStatusBar = booStatusBarState
''//Reset Status bar display setting

    Application.StatusBar = False
''//Return control of the Status bar to Excel

    Application.ScreenUpdating = True
''//Turn on screen updating

End Sub

'How to use. In each sub add
'Sub Main()
'StartSub:
'   On Error GoTo ErrCtrl
'   CallingSubName = SubName
'   SubName = "Main"
'   SwitchOff(True) 'Enable VBA_Performance module
'
'   'do your processing here
'
'
'ExitSub:
'    SwitchOff (False) 'Disable VBA_Performance module
'    SubName = CallingSubName
'    Exit Sub
'
'ErrCtrl:
'    ErrPrint Err.Number, Err.Description, SubName
'    Err.Clear
'    GoTo ExitSub
'
'End Sub

Sub UIHide()
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
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub UIShow()
    With Application
        .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",True)"
        .DisplayStatusBar = True
        .DisplayScrollBars = True
        .DisplayFormulaBar = True
    End With
    With ActiveWindow
        .DisplayHeadings = True
        .DisplayWorkbookTabs = True
        .DisplayRuler = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
End Sub


' Go to Tools -> References... and check "Microsoft Scripting Runtime"
' GUID{420B2830-E718-11CF-893D-00A0C9054228} to be able to use
' the FileSystemObject which has many useful features for handling files and folders

Public Sub SaveTextToFile(txt2wrt As String, FileName As String)

    ' The advantage of correctly typing fso as FileSystemObject is to make autocompletion
    ' (Intellisense) work, which helps you avoid typos and lets you discover other useful
    ' methods of the FileSystemObject
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    
    Dim fileStream As TextStream
    
    Dim filePath As String
    filePath = Application.ActiveWorkbook.Path & Application.PathSeparator & "Export" & Application.PathSeparator
    
    ' Here is a method of the FileSystemObject that checks if a folder exists
    If Not fso.FolderExists(filePath) Then fso.CreateFolder filePath '<~~~ doesn't exist, so create the folder
    
    filePath = filePath & FileName '<~~~ Adds file name to path
    
    ' Here is a method of the FileSystemObject that checks if a file exists
    If Not fso.FileExists(filePath) Then
        Set fileStream = fso.CreateTextFile(filePath) '<~~~ Here the actual file is created and opened for write access
    Else
        Set fileStream = fso.OpenTextFile(filePath, ForAppending) '<~~~ Here the file is opened for write access
    End If

    ' Write something to the file
    fileStream.WriteLine txt2wrt

    ' Close it, so it is not locked anymore
    fileStream.Close

    ' Explicitly setting objects to Nothing should not be necessary in most cases, but if
    ' you're writing macros for Microsoft Access, you may want to uncomment the following
    ' two lines (see https://stackoverflow.com/a/517202/2822719 for details):
    'Set fileStream = Nothing
    'Set fso = Nothing

End Sub
