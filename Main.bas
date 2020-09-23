Attribute VB_Name = "modMain"
Option Explicit


Sub Main()
    
    Dim fViewer As frmPwdView
    
    Set fViewer = New frmPwdView
    fViewer.Show
    Set fViewer = Nothing
    
End Sub
