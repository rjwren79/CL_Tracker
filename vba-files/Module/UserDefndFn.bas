Attribute VB_Name = "UserDefndFn"
'Finding the last Row of specified Sheet
Function fn_LastRow(ByVal SHT As Worksheet)

    Dim lastRow As Long
    lastRow = SHT.Cells.SpecialCells(xlLastCell).row
    lrow = SHT.Cells.SpecialCells(xlLastCell).row
    Do While Application.CountA(SHT.Rows(lrow)) = 0 And lrow <> 1
        lrow = lrow - 1
    Loop
    fn_LastRow = lrow
    Debug.Print "Function fn_LastRow was used"
End Function

Function IsNullOrEmpty(s As String) As Boolean

    If s = vbNullString Or s = Empty Or s = Null Then
        IsNullOrEmpty = True
    Else
        IsNullOrEmpty = False
    End If
'    Debug.Print "Function IsNullOrEmpty used"
End Function

Public Function nDate(xValue As String) As Variant


   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "nDate"
  
    If IsNullOrEmpty(xValue) Then
        nDate = Null
    Else
        nDate = CDate(xValue)
    End If
  
ExitFun:
    SubName = CallingSubName
    Exit Function
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    GoTo ExitFun
    
End Function

Public Function TransposeArray(myarray As Variant) As Variant
Dim X As Long
Dim Y As Long
Dim Xupper As Long
Dim Yupper As Long
Dim tempArray As Variant
    Xupper = UBound(myarray, 2)
    Yupper = UBound(myarray, 1)
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = myarray(Y, X)
        Next Y
    Next X
    TransposeArray = tempArray
End Function

Function RGB_To_vbHEX(ByVal rgb As String) As String

    rgb = CStr(rgb)
    'Replace(rgb, "RGB(", "")
    Debug.Print rgb
    'iRed As Integer, iGreen As Integer, iBlue As Integer
'    Dim sHex As String
'    sHex = "#" & VBA.Right$("00" & VBA.Hex(iBlue), 2) & VBA.Right$("00" & VBA.Hex(iGreen), 2) & VBA.Right$("00" & VBA.Hex(iRed), 2)
'    VBA_RGB_To_HEX = sHex
End Function

Function Concatenate(ParamArray Strings()) As String
    Concatenate = Join(Strings, vbNullString)
End Function

