VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPwdViewBand 
   BorderStyle     =   0  'None
   Caption         =   "Password Viewer"
   ClientHeight    =   5805
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   2745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgList 
      Left            =   960
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PwdViewBand.frx":0000
            Key             =   "find"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PwdViewBand.frx":015A
            Key             =   "settings"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PwdViewBand.frx":02B4
            Key             =   "login"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PwdViewBand.frx":040E
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   582
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnLogin"
            Object.ToolTipText     =   "Change login"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnOpen"
            Object.ToolTipText     =   "Open database"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnSettings"
            Object.ToolTipText     =   "Display settings"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnFind"
            Object.ToolTipText     =   "Find in password list"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   30
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   150
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwPwd 
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Resource"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblFind 
      Caption         =   "Find:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   75
      Width           =   375
   End
End
Attribute VB_Name = "frmPwdViewBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' PwdViewBands.frmPwdViewBand
' Form to be displayed in Explorer Band
'
'------------------------------------------------------------------------------
' The code in based on Explorer Band example from
' Eduardo Morcillo (edanmo@geocities.com)
' http://www.domaindlx.com/e_morcillo
' This project also uses type libs IObjectWithSite & Band interfaces and
' Storage and Property Sets interfaces from Eduardo Morcillo.
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit

' Explorer window reference
Public WithEvents IEWindow As InternetExplorer
Attribute IEWindow.VB_VarHelpID = -1

' Send message to a window, used to auto size listview columns
' based on contents
Private Declare Function SendMessage Lib "user32.dll" _
     Alias "SendMessageA" (ByVal hWnd As Long, _
     ByVal Msg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

Private msLastFile As String


Private Function CreateConnectionString(ByVal sFileName As String) As String
    
    If Len(sFileName) > 0 Then
        CreateConnectionString = "Provider=Microsoft.Jet.OLEdb.4.0;" & _
                    "Data Source=" & sFileName
    End If
    
End Function

Private Sub RefreshList()
    
    Dim oRs As ADODB.Recordset
    Dim oItem As MSComctlLib.ListItem
    Dim oSubItem As MSComctlLib.ListSubItem
    Dim sPwd As String
    Dim bClearText As Boolean
    
On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    DoEvents
    
    lvwPwd.ListItems.Clear
    bClearText = g_oApp.ShowClearTextPwd
    Set oRs = g_oApp.GetAllPwds
    If Not oRs Is Nothing Then
        With oRs
            While Not .EOF
                Set oItem = lvwPwd.ListItems.Add(Text:=!Resource.Value)
                oItem.Tag = !id.Value
                oItem.SubItems(1) = !UserID.Value
                sPwd = !Password.Value
                Set oSubItem = oItem.ListSubItems.Add
                oSubItem.Tag = sPwd
                If Not bClearText Then sPwd = String$(Len(!Password.Value), "*")
                oSubItem.Text = sPwd
                .MoveNext
            Wend
            .Close
        End With
    End If
    
errCleanup:
    If lvwPwd.ListItems.Count > 0 Then
        Set lvwPwd.SelectedItem = lvwPwd.ListItems(1)
    End If
    ResizeListViewColumns
    
    Set oRs = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    Screen.MousePointer = vbDefault
    MsgBox "Cannot refresh password list." & vbCrLf & vbCrLf & _
            "ERROR: " & Err.Number & ", " & Err.Description, _
            vbExclamation, App.ProductName
    Resume errCleanup
End Sub

Private Sub ResizeListViewColumns()
    
    Dim idx As Long
    Dim lngHwnd As Long
    
    lngHwnd = lvwPwd.hWnd
    
    For idx = 0 To lvwPwd.ListItems.Count - 1
        SendMessage lngHwnd, _
               LVM_SETCOLUMNWIDTH, _
               idx, _
               LVSCW_AUTOSIZE_USEHEADER
    Next idx
    
End Sub

Private Sub FindInList(ByVal sFind As String)
    
    Dim oItem As MSComctlLib.ListItem
    Dim idx As Long
    Dim lStart As Long
    
    If Len(sFind) > 0 Then
        With lvwPwd
            If Not .SelectedItem Is Nothing Then
                lStart = .SelectedItem.Index + 1
                If lStart > .ListItems.Count Then lStart = 1
            Else
                lStart = 1
            End If
            sFind = LCase$(sFind)
            
            For idx = lStart To .ListItems.Count
                Set oItem = .ListItems(idx)
                If InStr(LCase$(oItem.Text), sFind) > 0 Then
                    oItem.Selected = True
                    Exit For
                End If
            Next idx
        End With
    End If
            
End Sub


Private Sub Form_Load()
    
    Dim sConn As String
    
On Error GoTo errHandler
    ' Add child style to the window
    SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) Or WS_CHILD
    
    If g_oApp Is Nothing Then Set g_oApp = CreateObject("PasswordViewBus.Application")
    g_oApp.dologin
    RefreshList
    
    Exit Sub
    
errHandler:
    MsgBox "Cannot connect to password database." & vbCrLf & vbCrLf & _
            "ERROR: " & Err.Number & ", " & Err.Description, _
            vbExclamation, App.ProductName
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = vbMinimized Then
        If Me.Width < 860 Then Me.Width = 860
        If Me.Height < 800 Then Me.Height = 800
        lvwPwd.Move Me.ScaleLeft, Me.ScaleTop + tbrMenu.Height, _
                Me.ScaleWidth, Me.ScaleHeight - tbrMenu.Height - 60
        txtFind.Move tbrMenu.Buttons("btnFind").Left + tbrMenu.Buttons("btnFind").Width + 30, 30
    End If
End Sub


Private Sub lvwPwd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    With lvwPwd
        If .SortKey = (ColumnHeader.Index - 1) Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
    
End Sub

Private Sub lvwPwd_DblClick()
    Dim inp As IHTMLInputElement
    Dim el As IHTMLElement
    
On Error GoTo errHandler
    
    If Not lvwPwd.SelectedItem Is Nothing Then
        Clipboard.Clear
        Clipboard.SetText lvwPwd.SelectedItem.ListSubItems(2).Tag
        Set el = IEWindow.document.activeElement
        If el.tagName = "INPUT" Then
            Set inp = el
            If inp.Type = "password" Then
                inp.Value = lvwPwd.SelectedItem.ListSubItems(2).Tag
                Set el = IEWindow.document.All(el.sourceIndex - 1)
                If el.tagName = "INPUT" Then
                    Set inp = el
                    inp.Value = lvwPwd.SelectedItem.ListSubItems(1).Text
                End If
            Else
                inp.Value = lvwPwd.SelectedItem.ListSubItems(1).Text
                Set el = IEWindow.document.All(el.sourceIndex + 1)
                If el.tagName = "INPUT" Then
                    Set inp = el
                    If inp.Type = "password" Then
                        inp.Value = lvwPwd.SelectedItem.ListSubItems(2).Tag
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
errHandler:
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Dim bClearText As Boolean
    Dim oItem As MSComctlLib.ListItem
    Dim oSubItem As MSComctlLib.ListSubItem
    Dim sFile As String
    
On Error GoTo errHandler
    
    Select Case Button.Key
        Case "btnLogin"
            Me.MousePointer = vbHourglass
            If g_oApp.dologin Then RefreshList
            Me.MousePointer = vbDefault
            
        Case "btnOpen"
            With cdlg
                .CancelError = True
                .Filter = "All Files (*.*)|*.*|Database Files (*.mdb)|*.mdb"
                .FilterIndex = 2
                .DialogTitle = "Select password database"
                .InitDir = App.Path
                .Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
                .ShowOpen
                sFile = .FileName
            End With
            g_oApp.ConnectionString = CreateConnectionString(sFile)
            RefreshList
    
        Case "btnFind"
            FindInList txtFind.Text
            
        Case "btnSettings"
            bClearText = g_oApp.ShowClearTextPwd
            If g_oApp.ShowSettings Then
                If bClearText <> g_oApp.ShowClearTextPwd Then
                    bClearText = g_oApp.ShowClearTextPwd
                    For Each oItem In lvwPwd.ListItems
                        Set oSubItem = oItem.ListSubItems(2)
                        If bClearText Then
                            oSubItem.Text = oSubItem.Tag
                        Else
                            oSubItem.Text = String$(Len(oSubItem.Tag), "*")
                        End If
                    Next oItem
                End If
            End If
            
    End Select
    
    Exit Sub
errHandler:
    Me.MousePointer = vbDefault
End Sub

Private Sub txtFind_Change()
    tbrMenu.Buttons("btnFind").Enabled = (Len(txtFind.Text) > 0)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        DoEvents
        FindInList txtFind.Text
    End If
End Sub
