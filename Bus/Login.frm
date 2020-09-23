VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAutoLogin 
      Caption         =   "Log in automatically in the future"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   325
      Left            =   2760
      TabIndex        =   4
      Top             =   660
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   325
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblMsg 
      Caption         =   "Message to user"
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label lblPwd 
      Caption         =   "Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblUserID 
      Caption         =   "User ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' PasswordViewBus.frmLogin
' Login dialog, prompts user to enter Id and password, and returns them.
' Also allows to select Auto Login option and stores it in the registry
'
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit

Private mbCancelled As Boolean

Public Function DoLogin(ByRef sUID As String, _
                    ByRef sPwd As String) As Boolean
    
    Me.Caption = App.ProductName & " Login"
    lblMsg.Caption = "NOTE: Login Id and password are not stored anywhere " & _
                    "and all entries are ecnrypted so " & _
                    "if you forget login Id and password " & _
                    "you will not be able to get your entries."
    mbCancelled = True
    
    Me.Show vbModal
    
    DoLogin = Not mbCancelled
    If Not mbCancelled Then
        sUID = txtUserID.Text
        sPwd = txtPwd.Text
    End If
    Unload Me
End Function


Private Sub cmdCancel_Click()
    mbCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    SaveSetting App.Title, "Settings", "AutoLogin", chkAutoLogin.Value
    mbCancelled = False
    Me.Hide
End Sub

Private Sub txtPwd_Change()
    cmdOK.Enabled = (Len(txtUserID) > 0 And Len(txtPwd.Text) > 0)
End Sub

Private Sub txtUserID_Change()
    cmdOK.Enabled = (Len(txtUserID) > 0 And Len(txtPwd.Text) > 0)
End Sub
