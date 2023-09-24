Attribute VB_Name = "Module1"
Public Function Get_Field_Names()
StartSub:
   On Error GoTo ErrCtrl
   CallingSubName = SubName
   SubName = "Get_Field_Names"
    
    Dim rst As New ADODB.Recordset
    Dim fld As ADODB.Field
    Dim fldName As String
    Dim cString As String
    
    Dim qry As String
    qry = "SELECT * FROM EmpDatabase Where db_ID = " & 0

    ConnectDatabase
       
    rst.Open qry, DBCON, adOpenKeyset, adLockOptimistic
    
    On Error GoTo ConnErr
        
    Dim ctrl As Control
    Dim pos As String
    For Each fld In rst.Fields
            fldName = fld.Name
            cString = "rst.Fields(" & Chr(34) & fldName & Chr(34) & ").value = " & "." & fldName & ".Value"
            SaveTextToFile cString, "Save_Field_Names.txt"
    Next
    
    On Error GoTo ErrCtrl
    
    rst.Close
    CloseDatabase
       
ExitFun:
'    SwitchOff (True) 'On 'Disable VBA_Performance module
    SubName = CallingSubName
    Exit Function
    
ConnErr:
    Debug.Print fldName
    Resume Next
    
ErrCtrl:
    ErrPrint Err.Number, Err.Description, SubName
    Err.Clear
    CloseDatabase
    GoTo ExitFun
    
End Function
