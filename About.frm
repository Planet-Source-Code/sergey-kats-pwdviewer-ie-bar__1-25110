VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtThanks 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox txtRedist 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   780
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   325
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   120
      X2              =   3960
      Y1              =   1245
      Y2              =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   120
      X2              =   3960
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   3960
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   120
      X2              =   3960
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Build 2000.0.1"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Password Viewer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' PasswordView.frmAbout
' About window
'
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit

Private Sub Form_Load()
    lblTitle.Caption = App.ProductName
    lblVersion.Caption = "Build " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyright.Caption = "Copyright 2000, Sergey Kats. All rights reserved."
    txtRedist.Text = "You are free to ditribute, change, recompile any and all parts " & _
                    "of this application."
    txtThanks.Text = "Special thanks to " & _
                    "Edanmo VB (http://www.domaindlx.com/e_morcillo/default.asp) " & _
                    "for shell type libraries that made Explorer bar possible." & vbCrLf & _
                    "Some of cryptography functionality was developed with help of " & _
                    "VBPJ article by Dan Appleman from January 2001 issue."
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub
