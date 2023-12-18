Attribute VB_Name = "Access_Table_Functions"
Option Explicit
Public Function CreateTable(TableName As String) As Boolean

    Dim cnt As ADODB.Connection
    Dim strCreateTable As String
 
    On Error GoTo Err
 
    strCreateTable = "CREATE TABLE " & TableName
    
    Set cnt = New ADODB.Connection
    
    ConnectDatabase
    
    With cnt
        .Open DBCON
        .Execute strCreateTable
        .Close
    End With
    
    CloseDatabase
    
    If Err.Number = 0 Then
        CreateTable = True
        Debug.Print "Table '" & TableName & "' Created"
    End If

    Exit Function
    
Err:
        CreateTable = False
        Debug.Print "Error: " & Err.Description
        
End Function

Public Function TableExists(ByVal TableName As String) As Boolean
    
    'Function: Determine if table exists in an Access database
    'Arguments:strTablename:   Name of table to check
    Dim cnt As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim fld As ADODB.Field
    Dim strTableName As String
    
    Set cnt = New ADODB.Connection

    ConnectDatabase
    
    With cnt
        .Open DBCON
        Set rst = .OpenSchema(adSchemaTables)
        
       While Not rst.EOF
        If rst.Fields("TABLE_TYPE") = "TABLE" Then _
            strTableName = rst.Fields("TABLE_NAME")
            If strTableName = TableName Then
                TableExists = True
                GoTo ExitFun
            End If
        TableExists = False
        rst.MoveNext
        Wend
        
ExitFun:
        Set rst = Nothing
        .Close
    End With

    CloseDatabase
    
    Exit Function


End Function
Public Function FieldExists(ByVal TableName As String, FieldName As String) As Boolean
    
    'Function: Determine if fields exists within a table in an Access database
    'Arguments:strTablename:   Name of table to check
    Dim cnt As ADODB.Connection
    Dim fld As ADODB.Field
    Dim rst As ADODB.Recordset
    Dim strColumnName As String
    
    Set cnt = New ADODB.Connection
    
    ConnectDatabase
    
    With cnt
        .Open DBCON
        Set rst = .OpenSchema(adSchemaColumns)
        
        While Not rst.EOF
        If rst.Fields("TABLE_NAME") = TableName Then
            strColumnName = rst.Fields("COLUMN_NAME")
            If strColumnName = FieldName Then
                FieldExists = True
                GoTo ExitFun
            End If
        End If
        rst.MoveNext
        Wend
        FieldExists = False
        
ExitFun:
        Set rst = Nothing
        .Close
    End With

    CloseDatabase
    
    Exit Function


End Function

Public Function DeleteTable(TableName As String)

    If Not TableExists(TableName) Then
        Debug.Print "Table '" & TableName & "' Not Found!"
        Exit Function
    End If
        
    Dim cnt As ADODB.Connection
    Dim strDeleteTable As String
    strDeleteTable = "DROP TABLE [" & TableName & "];"
    
    Set cnt = New ADODB.Connection
    
    ConnectDatabase
    
    With cnt
        .Open DBCON
        .Execute strDeleteTable
        .Close
    End With
    
    CloseDatabase
    
    If Err.Number = 0 Then _
        Debug.Print "Table '" & TableName & "' was removed."
    
    Exit Function
        
Err:
        Debug.Print "Error: " & Err.Description
        Exit Function

End Function

Public Sub AddFieldToTable(TableName As String, FieldName As String, _
      FieldType As Long, FieldLen As Long, FieldAllowsNull As Boolean)

On Error GoTo Err

    Dim FieldText As String
    
    Select Case (FieldType)
        Case 0:
            FieldText = "Long"
        Case 1:
            FieldText = "text(" & FieldLen & ")"
        Case 2:
            FieldText = "bit"
        Case 3:
            FieldText = "datetime"
        Case 4:
            FieldText = "memo"
    
    End Select
    
    Dim Sql As String
    Sql = "ALTER TABLE " & TableName & " ADD COLUMN " & FieldName & " " & FieldText
    
    If FieldAllowsNull Then
       Sql = Sql & " NULL"
    Else
       Sql = Sql & " NOT NULL"
    End If

    If Not TableExists(TableName) Then
        Debug.Print "Table Not Found!"
        Exit Sub
    End If

    Dim cnt As ADODB.Connection
    Set cnt = New ADODB.Connection
    
    ConnectDatabase
    
    With cnt
        .Open DBCON
        .Execute Sql
        .Close
    End With
    
    CloseDatabase
    
    If Err.Number = 0 Then
        Debug.Print "Field '" & FieldName & "' was added to table '" & TableName & "'."
        Exit Sub
    End If
    
Err:
    If Err.Number = -2147217887 Then
        Debug.Print "Error: " & Err.Description
        Exit Sub
    End If
End Sub

Public Sub DelFieldOnTable(TableName As String, FieldName As String)
On Error GoTo Err
    
    Dim Sql As String
    Sql = "ALTER TABLE [" & TableName & "] DROP COLUMN [" & FieldName & "]"
    
    If Not TableExists(TableName) Then _
        Debug.Print "Table '" & TableName & "' Not Found!"

    If Not FieldExists(TableName, FieldName) Then _
        Debug.Print "Field '" & FieldName & "' was not found in table '" & TableName & "'."
    
    Dim cnt As ADODB.Connection
    Dim fld As ADODB.Field
    
    Set cnt = New ADODB.Connection
    
    ConnectDatabase
    
    With cnt
        .Open DBCON
        .Execute Sql
        .Close
    End With
    
    CloseDatabase
    
    If Err.Number = 0 Then _
        Debug.Print "Field '" & FieldName & "' was removed from table '" & TableName & "'."
    
    Exit Sub
        
Err:
        Debug.Print "Error: " & Err.Description
        Exit Sub
End Sub

