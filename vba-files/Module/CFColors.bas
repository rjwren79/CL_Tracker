Attribute VB_Name = "CFColors"
Function GetColors(Status As String) As Long

    If Status = "Granted" Then
        GetColors = rgb(56, 194, 56)  'Sets row color to dark washed chartreuse
    ElseIf Status = "Pending" Then
        GetColors = rgb(102, 189, 255) 'Sets Pending to lt. blue
    ElseIf Status = "PRdue" Then
        GetColors = rgb(192, 0, 0) 'Sets Denied row to guardsman red
    ElseIf Status = "Overdue" Then
        GetColors = rgb(255, 204, 204) 'Sets row color to pink
    ElseIf Status = "Expiring" Then
        GetColors = rgb(255, 204, 0) 'Sets row color to tangerine yellow
    End If

End Function


'        'eQIP
'        Cells(6, 19).Interior.Color = RGB(255, 255, 0)
'        'FP
'        Cells(7, 19).Interior.Color = RGB(112, 48, 160)
'        Cells(7, 19).Font.Color = vbWhite
'        'Needs Review
'        Cells(8, 19).Interior.Color = RGB(255, 0, 0)
'        'Pending BGC
'        Cells(9, 19).Interior.Color = RGB(255, 192, 0)
'        'Sec Briefs
'        Cells(10, 19).Interior.Color = RGB(102, 255, 153)
'        'CSR
'        Cells(11, 19).Interior.Color = RGB(221, 235, 247)
'        'Release
'        Cells(12, 19).Interior.Color = RGB(0, 0, 255)
'        'NDA
'        Cells(13, 19).Interior.Color = RGB(102, 0, 255)
'        'eQIP Term
'        Cells(14, 19).Interior.Color = RGB(0, 0, 0)
'        Cells(14, 19).Font.Color = vbRed
'        'Elgi Pending
'        Cells(15, 19).Interior.Color = RGB(224, 176, 134)
'        'PR Due
'        Cells(16, 19).Interior.Color = RGB(192, 0, 0)
'        Cells(16, 19).Font.Color = vbYellow
'        Cells(16, 19).Font.Bold = True
