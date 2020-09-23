VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' PasswordViewBus.clsDA
' Data access class
'
'   - ConnectionString  Returns/Sets connection string
'   - DeletePwd         Deletes password entry from db
'   - GetAllPwds        Returns recordset of all entries that match
'                       logged in user
'   - Password          Sets password of the currently logged in user
'   - SavePwd           Inserts new or updates existing entry
'   - SID               Returns/Sets SID of the currently logged in user
'   - UserId            Sets user Id of the currently logged in user
'
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit

' This is our encryption key
' It is only used to encrypt user name and password into SID
' which can be stored in the registry if user chooses to auto login
' The actual password entries are encrypted using login user name and password
Private Const SID_ENCRYPT_KEY As String = "lakfujgug4947rufj4j4fju4"

Private Const DELIM As String = "-"

Private msConn As String
Private msUID As String
Private msPwd As String


Public Property Let UserID(ByVal NewVal As String)
    msUID = NewVal
End Property

Public Property Let Password(ByVal NewVal As String)
    msPwd = NewVal
End Property

'
' Generates SID from user ID and password
' SID is encrypted using hard coded key
' SID is used only in case the user chooses to
' auto login, it is not used to access database
Public Property Get SID() As String
    
    Dim oCrypt As clsCrypto
    
    Set oCrypt = New clsCrypto
    SID = oCrypt.DoEncrypt(msUID & Chr$(0) & msPwd, SID_ENCRYPT_KEY)
    Set oCrypt = Nothing
    
End Property

'
' Decrypts SID and extract user ID and password
' which can be used to decrypt the password entries
' in the database
'
Public Property Let SID(ByVal NewVal As String)
    
    Dim oCrypt As clsCrypto
    Dim sSID As String
    Dim l As Long
    
    Set oCrypt = New clsCrypto
    sSID = oCrypt.DoDecrypt(NewVal, SID_ENCRYPT_KEY)
    Set oCrypt = Nothing
    
    l = InStr(sSID, Chr$(0))
    If l > 0 Then
        msUID = Left$(sSID, l - 1)
        msPwd = Mid$(sSID, l + 1)
    End If
    
End Property


Public Property Let ConnectionString(ByVal NewVal As String)
    msConn = NewVal
End Property

Public Property Get ConnectionString() As String
    ConnectionString = msConn
End Property

'
' Gets the list of all password entires and decrypts them
' using logged in user id and password
' It also comparing hash of user id and password that is stored in db also
' if hashes don't match that password entry is not returned
'
Public Function GetAllPwds() As ADODB.Recordset
    
    Dim oRs As ADODB.Recordset
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    Dim sUID As String
    Dim sPwd As String
    Dim sHash As String
    Dim oCrypt As clsCrypto
    Dim idx As Long
    Dim var As Variant
    
On Error GoTo errHandler
    
    sSQL = "select * from Passwords order by Resource"
    
    Set oConn = New ADODB.Connection
    oConn.Open msConn
    
    Set oRs = New ADODB.Recordset
    With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
        Set .ActiveConnection = Nothing
        oConn.Close
        Set oCrypt = New clsCrypto
        While Not .EOF
            ' Each password entry in the database contains
            ' site name, login name for that site, password for the site and
            ' hash of login name and password
            sUID = !UserID.Value
            sPwd = !Password.Value
            sHash = !Hash.Value & ""
            
            ' Login name, password and hash are stores as DELIM
            ' delimited string of ascii codes
            ' Convert from ASCII numbers to chars
            sUID = AsciiDelimitedToChar(sUID)
            sPwd = AsciiDelimitedToChar(sPwd)
            sHash = AsciiDelimitedToChar(sHash)
            
            ' Decrypt the password from DB using login password as a key
            sPwd = oCrypt.DoDecrypt(sPwd, msUID & msPwd)
            ' Decrypt the user name from DB using login password as a key
            sUID = oCrypt.DoDecrypt(sUID, msUID & msPwd)
            
            ' Get a hash of the decrypted user ID value and
            ' compare it to the origingal hash
            If sHash = oCrypt.DoHash(sUID & sPwd) Then
                !UserID.Value = sUID
                !Password.Value = sPwd
            Else
                ' Remove password entry if hashes don't match
                ' It's not deleted physically from db, just removed
                ' from the recordset
                .Delete
            End If
            .MoveNext
        Wend
        If .RecordCount > 0 Then .MoveFirst
    End With
    
errClenaup:
    Set oConn = Nothing
    
    Set GetAllPwds = oRs
    
    Exit Function
    
errHandler:
    MsgBox "Cannot open passwords database." & vbCrLf & _
            "Make sure database exists and connection string is correct." & vbCrLf & vbCrLf & _
            "ERROR: " & Err.Number & ", " & Err.Description, _
            vbExclamation, App.ProductName
    Set oRs = Nothing
    Resume errClenaup
End Function

'
' Saves new entry or updates existing entry
'
Public Function SavePwd(ByRef lID As Long, _
                    ByVal sResource As String, _
                    ByVal sUID As String, _
                    ByVal sPwd As String _
                    ) As Boolean
    
    Dim oConn As ADODB.Connection
    Dim oRs As ADODB.Recordset
    Dim sSQL As String
    Dim oCrypt As clsCrypto
    Dim sHash As String
    Dim idx As Long
    Dim sOut As String
    
On Error GoTo errHandler
    
    Set oCrypt = New clsCrypto
    With oCrypt
        ' Generate hash of new user name password pair
        sHash = .DoHash(sUID & sPwd)
        ' Encrypt new user name and password
        sUID = .DoEncrypt(sUID, msUID & msPwd)
        sPwd = .DoEncrypt(sPwd, msUID & msPwd)
    End With
    Set oCrypt = Nothing
    
    sUID = CharToAsciiDelimited(sUID)
    sPwd = CharToAsciiDelimited(sPwd)
    sHash = CharToAsciiDelimited(sHash)
    
    If lID > 0 Then
        sResource = Replace$(sResource, "'", "''")
        sSQL = "update Passwords set Resource='" & sResource & "'," & _
                "UserID='" & sUID & "'," & _
                "[Password]='" & sPwd & "'," & _
                "Hash='" & sHash & "' " & _
                "where ID=" & lID
    Else
        sSQL = "insert into Passwords(Resource,UserID,[Password],Hash) values (" & _
                "'" & sResource & "'," & _
                "'" & sUID & "'," & _
                "'" & sPwd & "'," & _
                "'" & sHash & "')"
    End If
    
    Set oConn = New ADODB.Connection
    oConn.Open msConn
    oConn.Execute sSQL
    If lID = 0 Then
        Set oRs = New ADODB.Recordset
        With oRs
            .CursorLocation = adUseClient
            .Open "select @@IDENTITY", oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            lID = .Fields(0).Value
            .Close
        End With
    End If
    SavePwd = True
    
errClenaup:
    oConn.Close
    Set oConn = Nothing
    
    Exit Function
    
errHandler:
    MsgBox "Cannot save passwords into database." & vbCrLf & _
            "Make sure database exists and connection string is correct." & vbCrLf & vbCrLf & _
            "ERROR: " & Err.Number & ", " & Err.Description, _
            vbExclamation, App.ProductName
    SavePwd = False
    Resume errClenaup
End Function

Public Function DeletePwd(ByVal lID As Long) As Boolean
    
    Dim oConn As ADODB.Connection
    Dim sSQL As String
    
On Error GoTo errHandler
    
    sSQL = "delete from Passwords where ID=" & lID
    
    Set oConn = New ADODB.Connection
    oConn.Open msConn
    oConn.Execute sSQL
    DeletePwd = True
    
errClenaup:
    oConn.Close
    Set oConn = Nothing
    
    Exit Function
    
errHandler:
    MsgBox "Cannot delete passwords from database." & vbCrLf & _
            "Make sure database exists and connection string is correct." & vbCrLf & vbCrLf & _
            "ERROR: " & Err.Number & ", " & Err.Description, _
            vbExclamation, App.ProductName
    DeletePwd = False
    Resume errClenaup
End Function

'
' Converts string of ascii codes delimited by DELIM
' into string of characters
'
Private Function AsciiDelimitedToChar(ByVal sIn As String) As String
    
    Dim var As Variant
    Dim sOut As String
    Dim idx As Long
    
    var = Split(sIn, DELIM)
    If IsArray(var) Then
        For idx = LBound(var) To UBound(var)
            sOut = sOut & Chr$(var(idx))
        Next idx
    End If
    AsciiDelimitedToChar = sOut
    
End Function

'
' Converts string to string of ascii codes delimited by DELIM
'
Private Function CharToAsciiDelimited(ByVal sIn As String) As String
    
    Dim idx As Long
    Dim sOut As String
    
    For idx = 1 To Len(sIn)
        sOut = sOut & Asc(Mid$(sIn, idx, 1)) & DELIM
    Next idx
    CharToAsciiDelimited = Left$(sOut, Len(sOut) - 1)
    
End Function

