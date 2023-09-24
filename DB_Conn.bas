Attribute VB_Name = "DB_Conn"
Option Explicit

Public gstrConnexString As String
Public DBLOC As String
Public DBNAM As String
Public DBCON As ADODB.Connection
Public Function FileExists(filePath As String) As Boolean

Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = Dir(filePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function
Public Function ConnectDatabase() As ADODB.Connection
    On Error GoTo ConnectionError
    
    'Identify DataBase location and name
    'DataBase Location
    'DBLOC = "\\radstor.ad.radford-aap.com\g4sgs\Alerts_DB" '(RFAAP H drive)
    'DBLOC = Application.ActiveWorkbook.Path '(Build Testing Home)
    'DBLOC = "C:\Users\richard.wren\Documents\GitHub\Alerts_Excel_Build" '(Build Testing Work)
    Call CheckFileExists
    ' DataBase File Name
    DBNAM = "\CLTDB"
    
    'Debug.Print DBLOC & DBNAM
    
    'Combine connection variables
    gstrConnexString = "Provider=Microsoft.ACE.OLEDB.16.0;" _
    & "Data Source=" & DBLOC & DBNAM & ";Jet OLEDB:Database Password=Z8b+T$+gdH"
    Application.StatusBar = "Connecting to external database..."
    Application.Cursor = xlWait
    
    'Check if connection is not already open
    If DBCON Is Nothing Then
        'Instantiate new database connection object
        Set DBCON = New ADODB.Connection
    'Otherwise return existing ADO connection object
    ElseIf DBCON.State = adStateOpen Then
        Set ConnectDatabase = DBCON
        Application.StatusBar = "Done"
        Application.Cursor = xlDefault
        Exit Function
    End If

    'Define the connection string
    DBCON.ConnectionString = gstrConnexString
    
    'Open the connection
    DBCON.Open
    Set ConnectDatabase = DBCON
    Application.StatusBar = "Done"
    Application.Cursor = xlDefault
    Exit Function
    
    'Error event
ConnectionError:
    MsgBox "Failed to connect to Database" & vbCrLf & Err.Description & " (" & Err.Number & ")"
    Set ConnectDatabase = Nothing
    Application.StatusBar = "Done"
    Application.Cursor = xlDefault
    
    'ActiveWorkbook.Close SaveChanges:=False
    
End Function

Public Function CloseDatabase()
    
    On Error Resume Next
    'Close the connection
    DBCON.Close
    Set DBCON = Nothing
    
    'Reset the error handler
    On Error GoTo 0
    
End Function

Public Sub CheckFileExists()
On Error GoTo ErrCtrl

    DBLOC = Application.ActiveWorkbook.Path & "\DataBase"
    
ExitSub:
    Exit Sub

ErrCtrl:
    Debug.Print "Useing Local File"

End Sub


