Attribute VB_Name = "OleModule"
Public Sub Copy_ole_Positions()

    'Dim wb1 As Excel.Workbook
    Dim wb2 As Excel.Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim ole As OLEObject
    Dim oleArr As Variant
    Dim oleName As String
    Dim oleLeft As Integer
    Dim oleTop As Integer
    Dim oleHeight As Integer
    Dim oleWidth As Integer
    
    'Set wb1 =
    Set ws1 = ThisWorkbook.Worksheets("Alerts")
    Set wb2 = Excel.Workbooks("09052023.xlsm")
    Set ws2 = wb2.Worksheets("Alerts")
    
    For Each ole In ws2.OLEObjects
        oleArr = OLE_Position(ole)
        oleName = CStr(oleArr(0))
        oleLeft = CInt(oleArr(1))
        oleTop = CInt(oleArr(2))
        oleHeight = CInt(oleArr(3))
        oleWidth = CInt(oleArr(4))
        With ws1.OLEObjects(oleName)
            .Left = oleLeft
            .Top = oleTop
            .Height = oleHeight
            .Width = oleWidth
        End With
    Next

End Sub

Public Sub Set_ole_Positions()
    
    'Dim wb1 As Excel.Workbook
    Dim ws1 As Worksheet
    Dim ole As OLEObject
    Dim oleArr As Variant
    Dim oleName As String
    Dim oleLeft As Double
    Dim oleTop As Double
    Dim oleHeight As Double
    Dim oleWidth As Double
    Dim oleRight As Double
    
    
    'Set wb1 =
    Set ws1 = ThisWorkbook.Worksheets("Alerts")
    
    For Each ole In ws1.OLEObjects
        oleArr = OLE_Position(ole)
        oleName = CStr(oleArr(0))
        oleLeft = oleArr(1)
        oleTop = oleArr(2)
        oleHeight = oleArr(3)
        oleWidth = oleArr(4)
'        With ws1.OLEObjects(oleName)
'            .Left = oleLeft
'            .Top = oleTop
'            .Height = oleHeight
'            .Width = oleWidth
'        End With
        oleRight = oleLeft + oleWidth
        Debug.Print "oleName: " & oleName & vbCrLf _
                  & ".Left:   " & oleLeft & vbCrLf _
                  & ".Right:  " & oleRight & vbCrLf _
                  & ".Top:    " & oleTop & vbCrLf _
                  & ".Height: " & oleHeight & vbCrLf _
                  & ".Width:  " & oleWidth
    Next
    
    
End Sub
Public Function OLE_Position(ole As OLEObject) As Variant

    Dim oleName As String
    Dim oleLeft As String
    Dim oleTop As String
    Dim oleHeight As Integer
    Dim oleWidth As Integer

    With ole
        oleName = .Name
        oleLeft = .Left
        oleTop = .Top
        oleHeight = .Height
        oleWidth = .Width
    End With
    
    OLE_Position = Array(oleName, oleLeft, oleTop, oleHeight, oleWidth)

End Function

Sub addshapetocell(cl As Variant)

Dim clLeft As Double
Dim clTop As Double
Dim clWidth As Double
Dim clHeight As Double

'Dim cl As Range
'Dim strShape As String

Set cl = Range("D3")  '<-- Range("D3")
clLeft = cl.Left
clTop = cl.Top
clHeight = cl.Height
clWidth = cl.Width
clRight = clLeft + clWidth


'Set shpOval = ActiveSheet.Shapes.AddShape(msoShapeOval, clLeft, clTop, 4, 10)

Debug.Print "Shape .Left:   " & clLeft & vbCrLf _
          & "      .Right:  " & clRight & vbCrLf _
          & "      .Top:    " & clTop & vbCrLf _
          & "      .Height: " & clHeight & vbCrLf _
          & "      .Width:  " & clWidth

End Sub
