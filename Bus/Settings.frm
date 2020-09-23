VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtConn 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CheckBox chkAutoLogin 
      Caption         =   "Log in Automatically in the Future"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.CheckBox chkViewPwd 
      Caption         =   "View Passwords in Clear Text"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   325
      Left            =   3360
      TabIndex        =   4
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   325
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblConn 
      Caption         =   "Database Connection:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' PasswordViewBus.frmSettings
' Settings window, retrieves from and stores all settings in the registry
'
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit


Private mbCancelled As Boolean

Public Function ShowSettings() As Boolean
    
    Me.Caption = App.ProductName & " Settings"
    chkViewPwd.Value = GetSetting(App.Title, "Settings", "ViewClearTextPwd", 0)
    chkAutoLogin.Value = GetSetting(App.Title, "Settings", "AutoLogin", 0)
    txtConn.Text = GetSetting(App.Title, "Settings", "ConnString", "")
    
    mbCancelled = True
    cmdOK.Enabled = False
    
    Me.Show vbModal
    
    ShowSettings = Not mbCancelled
    If Not mbCancelled Then
        SaveSetting App.Title, "Settings", "ViewClearTextPwd", chkViewPwd.Value
        SaveSetting App.Title, "Settings", "AutoLogin", chkAutoLogin.Value
        SaveSetting App.Title, "Settings", "ConnString", txtConn.Text
    End If
    Unload Me
End Function


Private Sub cmdCancel_Click()
    mbCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mbCancelled = False
    Me.Hide
End Sub


Private Sub chkViewPwd_Click()
    cmdOK.Enabled = True
End Sub

Private Sub chkAutoLogin_Click()
    cmdOK.Enabled = True
End Sub

Private Sub txtConn_Change()
    cmdOK.Enabled = True
End Sub
