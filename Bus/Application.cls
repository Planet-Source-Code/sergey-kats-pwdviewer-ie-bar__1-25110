VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
' PasswordViewBus.Application
' Main class of PasswordView component. It exposes all functionality of the
' password viewer.
'
'   - ConnectionString  Returns/Sets connection string to the database
'   - DeletePwd         Delete password entry
'   - DoLogin           Logs user in
'   - GetAllPwds        Returns recordset of all entries
'   - SavePwd           Update/Inserts password entry
'   - ShowClearTextPwd  Returns/Sets flag to show clear text password in UI
'   - ShowSettings      Show settings window
'
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit


Private moDA As clsDA
Private mbShowClearTextPwd As Boolean


Private Sub Class_Initialize()
    
    Set moDA = New clsDA
    moDA.ConnectionString = GetSetting(App.Title, "Settings", "ConnString", "")
    mbShowClearTextPwd = GetSetting(App.Title, "Settings", "ViewClearTextPwd", 0)
    
End Sub


Public Function DoLogin(Optional ByVal sUID As String = "", _
                        Optional ByVal sPwd As String = "" _
                        ) As Boolean
    
    Dim fLogin As frmLogin
    Dim bResult As Boolean
    Dim sSID As String
    
    ' First check if AutoLogin flag is set
    If GetSetting(App.Title, "Settings", "AutoLogin", 0) = 1 Then
        ' Try to get SID from the registry
        sSID = GetSetting(App.Title, "Settings", "SID", "")
        If Len(sSID) > 0 Then
            moDA.SID = sSID
            bResult = True
        End If
    Else
        ' If user Id and password are passed in don't display
        ' login dialog
        If Len(sUID) > 0 And Len(sPwd) > 0 Then
            bResult = True
        Else
            ' Login dialog prompts the user for name/password
            ' Also it allows them to choose to login automatically
            ' in the future, that flag will be saved to the registry
            Set fLogin = New frmLogin
            bResult = fLogin.DoLogin(sUID, sPwd)
            Unload fLogin
            Set fLogin = Nothing
        End If
        
        If bResult Then
            With moDA
                ' Save user Id and password, DataAccess class will need
                ' them to decrypt the password list
                .UserID = sUID
                .Password = sPwd
                ' Check if user chose to auto login in the future
                ' Save SID to the registry
                If GetSetting(App.Title, "Settings", "AutoLogin", 0) = 1 Then _
                    SaveSetting App.Title, "Settings", "SID", .SID
            End With
        End If
        
    End If
    
    DoLogin = bResult
    
End Function

Public Function GetAllPwds() As ADODB.Recordset
    Set GetAllPwds = moDA.GetAllPwds
End Function

Public Function SavePwd(ByRef lID As Long, _
                    ByVal sResource As String, _
                    ByVal sUID As String, _
                    ByVal sPwd As String _
                    ) As Boolean
    SavePwd = moDA.SavePwd(lID, sResource, sUID, sPwd)
End Function

Public Function DeletePwd(ByVal lID As Long) As Boolean
    DeletePwd = moDA.DeletePwd(lID)
End Function


Public Property Let ConnectionString(ByVal NewVal As String)
    moDA.ConnectionString = NewVal
End Property

Public Property Get ConnectionString() As String
    ConnectionString = moDA.ConnectionString
End Property


Public Property Let ShowClearTextPwd(ByVal NewVal As Boolean)
    mbShowClearTextPwd = NewVal
End Property

Public Property Get ShowClearTextPwd() As Boolean
    ShowClearTextPwd = mbShowClearTextPwd
End Property


Public Function ShowSettings() As Boolean
    
    Dim frmX As frmSettings
    Dim bClearText As Boolean
    
    Set frmX = New frmSettings
    ShowSettings = frmX.ShowSettings
    Unload frmX
    Set frmX = Nothing
    
    If ShowSettings Then
        mbShowClearTextPwd = GetSetting(App.Title, "Settings", "ViewClearTextPwd", 0)
        ConnectionString = GetSetting(App.Title, "Settings", "ConnString", ConnectionString)
    End If
    
End Function


Private Sub Class_Terminate()
    SaveSetting App.Title, "Settings", "ConnString", moDA.ConnectionString
    SaveSetting App.Title, "Settings", "ViewClearTextPwd", Abs(CLng(mbShowClearTextPwd))
End Sub
