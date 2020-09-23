VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPwdView 
   Caption         =   "Password View - Store all your passwords in one place"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5730
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
   ScaleHeight     =   3600
   ScaleWidth      =   5730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFind 
      Caption         =   "8"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   5040
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwPwd 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   300
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
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
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogin 
         Caption         =   "Login..."
      End
      Begin VB.Menu mnuDelim4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Database..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Changes"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuDelim1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAddPwd 
         Caption         =   "Add &New Entry..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuChangePwd 
         Caption         =   "Change &Entry..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuDelPwd 
         Caption         =   "&Delete Entry"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDelim3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuDelim2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmPwdView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' PasswordView.frmPwdView
' Main form of the application, displays password list.
'
'------------------------------------------------------------------------------
' This code is provided "AS IS" without warranty of any kind.
' You are free to use and abuse this code any way you want.
' Do not distribute the original source and binaries.
' Copyright 2001, Sergey Kats. All rights reserved.
'------------------------------------------------------------------------------
Option Explicit

'
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
Private moApp As PasswordViewBus.Application

'
' Generate Access connection string from path to physical location of the mdb
'
Private Function CreateConnectionString(ByVal sFileName As String) As String
    
    If Len(sFileName) > 0 Then
        CreateConnectionString = "Provider=Microsoft.Jet.OLEdb.4.0;" & _
                    "Data Source=" & sFileName
    End If
    
End Function

'
' Refresh password list, gets recordset of all passwords
' and adds them to the list, replaces passwords with * if
' ShowClearTextPwd is not set.
'
Private Sub RefreshList()
    
    Dim oRs As ADODB.Recordset
    Dim oItem As MSComctlLib.ListItem
    Dim oSubItem As MSComctlLib.ListSubItem
    Dim sPwd As String
    Dim bClearText As Boolean
    
On Error GoTo errHandler
    Me.MousePointer = vbHourglass
    DoEvents
    
    lvwPwd.ListItems.Clear
    bClearText = moApp.ShowClearTextPwd
    Set oRs = moApp.GetAllPwds
    If Not oRs Is Nothing Then
        With oRs
            While Not .EOF
                Set oItem = lvwPwd.ListItems.Add(Text:=!Resource.Value)
                oItem.Tag = !ID.Value
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
        mnuCopy.Enabled = True
        mnuChangePwd.Enabled = True
        mnuDelPwd.Enabled = True
    Else
        mnuCopy.Enabled = False
        mnuChangePwd.Enabled = False
        mnuDelPwd.Enabled = False
    End If
    ResizeListViewColumns
    
    Set oRs = Nothing
    
    Me.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    MsgBox "Cannot refresh password list." & vbCrLf & vbCrLf & _
            "ERROR: " & Err.Number & ", " & Err.Description, _
            vbExclamation, App.ProductName
    Resume errCleanup
End Sub

'
' Resizes all list view columns
'
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

'
' Initializes PasswordView object
'
Private Sub Form_Load()
    
    Set moApp = CreateObject("PasswordViewBus.Application")
    Me.Visible = True
    DoEvents
    If moApp.DoLogin Then RefreshList
    
End Sub

Private Sub Form_Resize()
    If Not Me.WindowState = vbMinimized Then
        If Me.Width < 2685 Then Me.Width = 2700
        If Me.Height < 2115 Then Me.Height = 2115
        lvwPwd.Move Me.ScaleLeft, Me.ScaleTop + 300, _
                Me.ScaleWidth, Me.ScaleHeight - 300
    End If
End Sub

'
' Prompts to save changes if any to the database before unloading the form
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lResult As Long
    
    If mnuSave.Enabled Then
        lResult = MsgBox("Do you want to save changes before exiting the application?", _
                    vbQuestion + vbYesNoCancel, App.ProductName)
        If lResult = vbYes Then
            mnuSave_Click
        ElseIf (lResult = vbCancel) And _
                (UnloadMode <> vbAppWindows) And _
                (UnloadMode <> vbAppTaskManager) Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moApp = Nothing
End Sub

'
' Sorts list view when column heading is clicked
'
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

'
' Displays edit form when entry in the list is double clicked
'
Private Sub lvwPwd_DblClick()
    mnuChangePwd_Click
End Sub

Private Sub lvwPwd_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        mnuFind.Visible = False
        mnuDelim2.Visible = False
        If lvwPwd.HitTest(x, y) Is Nothing Then
           mnuChangePwd.Visible = False
        End If
        Me.PopupMenu mnuEdit, vbPopupMenuRightButton, x + lvwPwd.Left, y + lvwPwd.Top
        
        mnuFind.Visible = True
        mnuDelim2.Visible = True
        mnuChangePwd.Visible = True
    End If
End Sub

'
' Displays edit form and adds new entry to the list
' New entries are marked in red until they are saved
' to the database
'
Private Sub mnuAddPwd_Click()
    
    Dim frmX As frmEdit
    Dim oItem As MSComctlLib.ListItem
    Dim oSubItem As MSComctlLib.ListSubItem
    Dim sResource As String
    Dim sUID As String
    Dim sPwd As String
    
    Set frmX = New frmEdit
    frmX.ShowClearTextPwd = moApp.ShowClearTextPwd
    If frmX.AddRecord(sResource, sUID, sPwd) Then
        Set oItem = lvwPwd.ListItems.Add(Text:=sResource)
        With oItem
            .Tag = 0
            .ForeColor = vbRed
            Set oSubItem = .ListSubItems.Add(Text:=sUID)
            oSubItem.ForeColor = vbRed
            Set oSubItem = .ListSubItems.Add
            oSubItem.Tag = sPwd
            If Not moApp.ShowClearTextPwd Then sPwd = String$(Len(sPwd), "*")
            oSubItem.Text = sPwd
            oSubItem.ForeColor = vbRed
        End With
        mnuSave.Enabled = True
        mnuCopy.Enabled = True
        mnuChangePwd.Enabled = True
        mnuDelPwd.Enabled = True
        ResizeListViewColumns
    End If
    Set frmX = Nothing
    
End Sub

'
' Displays edit form and updates password entry
' Changed entries are shown in red until they are saved
' to the database
'
Private Sub mnuChangePwd_Click()
    
    Dim frmX As frmEdit
    Dim oSubItem As MSComctlLib.ListSubItem
    Dim sResource As String
    Dim sUID As String
    Dim sPwd As String
    
    If Not lvwPwd.SelectedItem Is Nothing Then
        Set frmX = New frmEdit
        frmX.ShowClearTextPwd = moApp.ShowClearTextPwd
        With lvwPwd.SelectedItem
            sResource = .Text
            sUID = .SubItems(1)
            sPwd = .SubItems(2)
            If frmX.EditRecord(sResource, sUID, sPwd) Then
                .ForeColor = vbRed
                .Text = sResource
                Set oSubItem = .ListSubItems(1)
                oSubItem.Text = sUID
                oSubItem.ForeColor = vbRed
                Set oSubItem = .ListSubItems(2)
                oSubItem.Tag = sPwd
                If Not moApp.ShowClearTextPwd Then sPwd = String$(Len(sPwd), "*")
                oSubItem.Text = sPwd
                oSubItem.ForeColor = vbRed
                mnuSave.Enabled = True
                ResizeListViewColumns
            End If
        End With
        Set frmX = Nothing
    End If
    
End Sub

'
' Marks password entry for deletion, this function doesn't
' actually deletes anything, it changes fore color of list items
' and they will be deleted when changes are saved to the db
'
Private Sub mnuDelPwd_Click()
    
    Dim oSubItem As MSComctlLib.ListSubItem
    
    If Not lvwPwd.SelectedItem Is Nothing Then
        lvwPwd.SelectedItem.ForeColor = vbMenuBar
        For Each oSubItem In lvwPwd.SelectedItem.ListSubItems
            oSubItem.ForeColor = vbMenuBar
        Next oSubItem
    End If
End Sub

'
' Copies password to the clipboard in clear text
'
Private Sub mnuCopy_Click()
    
    If Not lvwPwd.SelectedItem Is Nothing Then
        Clipboard.Clear
        Clipboard.SetText lvwPwd.SelectedItem.ListSubItems(2).Tag
    End If
    
End Sub

Private Sub mnuFind_Click()
    txtFind.SetFocus
End Sub

Private Sub mnuLogin_Click()
    
    Me.MousePointer = vbHourglass
    If moApp.DoLogin Then RefreshList
    Me.MousePointer = vbDefault
    
End Sub

'
' Prompts to select new Access mdb file, creates connection to it
' and loads list of passwords from new db
'
Private Sub mnuOpen_Click()
    
On Error GoTo errHandler
    
    With cdlg
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|Database Files (*.mdb)|*.mdb"
        .FilterIndex = 2
        .DialogTitle = "Select password database"
        If Len(msLastFile) = 0 Then
            .InitDir = App.Path
        Else
            .InitDir = msLastFile
        End If
        .Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
        .ShowOpen
        msLastFile = .FileName
    End With
    
    moApp.ConnectionString = CreateConnectionString(msLastFile)
    RefreshList
    
    Exit Sub
    
errHandler:
    ' User probably cancelled open dialog
End Sub

'
' Saves all changes to the database
' Inserts/updates entries marked red and deletes entries marked
' menu bar system color
'
Private Sub mnuSave_Click()
    
    Dim oItem As MSComctlLib.ListItem
    Dim oSubItem As MSComctlLib.ListSubItem
    Dim lID As Long
    Dim idx As Long
    Dim bNotSaved As Boolean
    
    Screen.MousePointer = vbHourglass
    DoEvents
    bNotSaved = False
    
    For idx = lvwPwd.ListItems.Count To 1 Step -1
        Set oItem = lvwPwd.ListItems(idx)
        With oItem
            If .ForeColor = vbRed Then
                lID = Val(.Tag)
                If moApp.SavePwd(lID, .Text, .SubItems(1), .ListSubItems(2).Tag) Then
                    .Tag = lID
                    .ForeColor = vbWindowText
                    Set oSubItem = .ListSubItems(1)
                    oSubItem.ForeColor = vbWindowText
                    Set oSubItem = .ListSubItems(2)
                    oSubItem.ForeColor = vbWindowText
                Else
                    bNotSaved = True
                End If
            ElseIf (.ForeColor = vbMenuBar) Then
                If .Tag > 0 Then
                    lID = Val(.Tag)
                    If moApp.DeletePwd(lID) Then
                        lvwPwd.ListItems.Remove .Index
                    Else
                        bNotSaved = True
                    End If
                Else
                    lvwPwd.ListItems.Remove .Index
                End If
            End If
        End With
    Next idx
    
    mnuSave.Enabled = bNotSaved
    
    Screen.MousePointer = vbDefault
    DoEvents
    
End Sub

'
' Shows settings form and updates the view accordingly
'
Private Sub mnuSettings_Click()
    
    Dim bClearText As Boolean
    Dim oItem As MSComctlLib.ListItem
    Dim oSubItem As MSComctlLib.ListSubItem
    
    bClearText = moApp.ShowClearTextPwd
    If moApp.showsettings Then
        If bClearText <> moApp.ShowClearTextPwd Then
            bClearText = moApp.ShowClearTextPwd
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
    
End Sub

'
' Shows about window
'
Private Sub mnuAbout_Click()
    
    Dim frmX As frmAbout
    
    Set frmX = New frmAbout
    frmX.Show vbModal
    
    Unload frmX
    Set frmX = Nothing
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

'
' Searches the list to find matches, it only
' looks in site names
'
Private Sub cmdFind_Click()
    
    Dim oItem As MSComctlLib.ListItem
    Dim idx As Long
    Dim lStart As Long
    Dim sFind As String
    
    If Len(txtFind.Text) > 0 Then
        With lvwPwd
            If Not .SelectedItem Is Nothing Then
                lStart = .SelectedItem.Index + 1
                If lStart > .ListItems.Count Then lStart = 1
            Else
                lStart = 1
            End If
            sFind = LCase$(txtFind.Text)
            
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

Private Sub txtFind_Change()
    cmdFind.Enabled = (Len(txtFind.Text) > 0)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        DoEvents
        cmdFind_Click
    End If
End Sub
