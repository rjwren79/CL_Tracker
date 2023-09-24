Attribute VB_Name = "SystemUser"
Option Explicit

Function GetUserName()

    Dim UserName As String
    Dim gun As String
    Dim pis As Integer
    
    gun = Environ$("UserName")
    pis = InStr(gun, ".")
    
    If pis <> 0 Then
'        GetUserName = StrConv(Left(gun, 1) & ". " & Right(gun, Len(gun) - InStr(gun, ".")), vbProperCase) 'F. Last
'        GetUserName = StrConv(Replace(gun, ".", " "), vbProperCase) 'First Last
        GetUserName = StrConv(Left(gun, InStr(gun, ".") - 1) & "!", vbProperCase) 'First name
    Else
        GetUserName = gun
    End If
    Exit Function

End Function
