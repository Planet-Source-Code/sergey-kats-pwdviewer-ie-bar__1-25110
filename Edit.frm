VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit password record"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Edit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   325
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   325
      Left            =   3480
      TabIndex        =   5
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox txtResource 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblPwd 
      Caption         =   "Confirm Password:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblPwd 
      Caption         =   "Password:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblResource 
      Caption         =   "Resource:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblUserID 
      Caption         =   "User ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' PasswordView.frmEdit
' Edit or add new password entry
'
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit

Private mbCancelled As Boolean
Private mbShowClearTextPwd As String

'
' Set this propperty before calling other functions
' it will either mask or unmask passwords
'
Public Property Let ShowClearTextPwd(ByVal NewVal As Boolean)
    mbShowClearTextPwd = NewVal
End Property

'
' Initialize the form
'
Private Sub Form_Load()
    If mbShowClearTextPwd Then
        txtPwd(0).PasswordChar = ""
        txtPwd(1).PasswordChar = ""
    Else
        txtPwd(0).PasswordChar = "*"
        txtPwd(1).PasswordChar = "*"
    End If
End Sub

'
' Call this function to edit entry
'
Public Function EditRecord(ByRef sResource As String, _
                    ByRef sUID As String, _
                    ByRef sPwd As String) As Boolean
    txtResource.Text = sResource
    txtUserID.Text = sUID
    txtPwd(0).Text = sPwd
    txtPwd(1).Text = sPwd
    mbCancelled = True
    CheckData
    
    Me.Caption = "Edit Password Record"
    Me.Show vbModal
    
    EditRecord = Not mbCancelled
    If Not mbCancelled Then
        sResource = txtResource.Text
        sUID = txtUserID.Text
        sPwd = txtPwd(0).Text
    End If
    Unload Me
End Function

'
' Add new entry
'
Public Function AddRecord(ByRef sResource As String, _
                    ByRef sUID As String, _
                    ByRef sPwd As String) As Boolean
    txtResource.Text = ""
    txtUserID.Text = ""
    txtPwd(0).Text = ""
    txtPwd(1).Text = ""
    cmdOK.Enabled = False
    mbCancelled = True
    
    Me.Caption = "Add New Password"
    Me.Show vbModal
    
    AddRecord = Not mbCancelled
    If Not mbCancelled Then
        sResource = txtResource.Text
        sUID = txtUserID.Text
        sPwd = txtPwd(0).Text
    End If
    Unload Me
End Function

'
' Make sure all required fields are filled in before
' enabling OK button
'
Private Sub CheckData()
    If Len(txtUserID.Text) > 0 And _
            Len(txtResource.Text) > 0 And _
            Len(txtPwd(0).Text) > 0 And _
            Len(txtPwd(1).Text) > 0 And _
            (txtPwd(0).Text = txtPwd(1).Text) Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    mbCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    mbCancelled = False
    Me.Hide
End Sub

Private Sub txtPwd_Change(Index As Integer)
    CheckData
End Sub

Private Sub txtResource_Change()
    CheckData
End Sub

Private Sub txtUserID_Change()
    CheckData
End Sub
