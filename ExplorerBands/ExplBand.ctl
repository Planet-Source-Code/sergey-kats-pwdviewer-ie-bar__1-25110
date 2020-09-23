VERSION 5.00
Begin VB.UserControl ExplBand 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ExplBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
' PwdViewBands.ExplBand
'
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

Implements IOWS.IOleWindow
Implements IOWS.IDockingWindow
Implements IOWS.IDeskBand

Implements IOWS.IObjectWithSite

Implements IOWS.IPersist
Implements IOWS.IPersistStream

' Band site object
Dim m_Site As IOWS.IUnknown

' Band window
Dim m_Band As frmPwdViewBand

Private Sub IDeskBand_CloseDW(ByVal dwReserved As Long)

   ' Call IDockingWindow implementation
   IDockingWindow_CloseDW dwReserved
   
End Sub

Private Sub IDeskBand_ContextSensitiveHelp(ByVal fEnterMode As Long)

   ' Not implemented
   
   Err.Raise E_NOTIMPL
   
End Sub

Private Sub IDeskBand_GetBandInfo(ByVal dwBandID As Long, ByVal dwViewMode As IOWS.GetBandInfo_ViewModes, pdbi As IOWS.DESKBANDINFO)
Dim sTitle As String
   
   On Error Resume Next
     
   With pdbi
      
      If (.dwMask And DBIM_MINSIZE) = DBIM_MINSIZE Then
         ' Set minimum size
         .ptMinSize.x = 100
         .ptMinSize.y = 22
      End If
      
      If (.dwMask And DBIM_MAXSIZE) = DBIM_MAXSIZE Then
         ' Set maximum size
         .ptMaxSize.y = -1
         .ptMaxSize.x = -1
      End If
      
      If (.dwMask And DBIM_ACTUAL) = DBIM_ACTUAL Then
         ' Set ideal size
         .ptActual.x = 100
         .ptActual.y = 100
      End If
      
      If (.dwMask And DBIM_INTEGRAL) = DBIM_INTEGRAL Then
         .ptIntegral.x = 1
         .ptIntegral.y = 1
      End If
      
      If (.dwMask And DBIM_TITLE) = DBIM_TITLE Then
         
         ' Set band title
         sTitle = m_Band.Caption & vbNullChar
         CopyMemory .wszTitle(0), ByVal StrPtr(sTitle), LenB(sTitle)
         
      End If
      
      If (.dwMask And DBIM_BKCOLOR) = DBIM_BKCOLOR Then
         .crBkgnd = GetSysColor(COLOR_BTNFACE)
      End If
      
      If (.dwMask And DBIM_MODEFLAGS) = DBIM_MODEFLAGS Then
         ' Set flags
         .dwModeFlags = DBIMF_VARIABLEHEIGHT Or DBIMF_BKCOLOR
      End If

   End With
   
End Sub

Private Function IDeskBand_GetWindow() As Long

   ' Call IDockingWindow implementation
   
   IDeskBand_GetWindow = IDockingWindow_GetWindow
   
End Function

Private Sub IDeskBand_ResizeBorderDW(prcBorder As IOWS.RECT, ByVal punkToolbarSite As Long, ByVal fReserved As Long)

   ' Not implemented
   
   Err.Raise E_NOTIMPL
   
End Sub

Private Sub IDeskBand_ShowDW(ByVal fShow As Long)

   ' Call IDockingWindow implementation
   IDockingWindow_ShowDW fShow
   
End Sub

Private Sub IDockingWindow_CloseDW(ByVal dwReserved As Long)
   
   On Error Resume Next

   ' Hide the UserControl
   Set m_Band = Nothing
   
End Sub

Private Sub IDockingWindow_ContextSensitiveHelp(ByVal fEnterMode As Long)

   ' Not implemented
   
   Err.Raise E_NOTIMPL

End Sub

Private Function IDockingWindow_GetWindow() As Long

   On Error Resume Next
   
   ' Call IOleWindow implementation
   
   IDockingWindow_GetWindow = IOleWindow_GetWindow
   
End Function

Private Sub IDockingWindow_ResizeBorderDW(prcBorder As IOWS.RECT, ByVal punkToolbarSite As Long, ByVal fReserved As Long)

   ' Not implemented
   
   Err.Raise E_NOTIMPL
   
End Sub

Private Sub IDockingWindow_ShowDW(ByVal fShow As Long)
   
   On Error Resume Next

   ' Show/Hide the window
   If fShow Then
      ShowWindow m_Band.hWnd, SW_SHOW
   Else
      ShowWindow m_Band.hWnd, SW_HIDE
   End If
   
End Sub

Private Sub IOleWindow_ContextSensitiveHelp(ByVal fEnterMode As Long)

   Err.Raise E_NOTIMPL
   
End Sub

Private Function IObjectWithSite_GetSite(riid As IOWS.UUID) As stdole.IUnknown

   ' Get the requested interface
   Set IObjectWithSite_GetSite = m_Site.QueryInterface(riid)
   
End Function

Private Sub IObjectWithSite_SetSite(ByVal pUnkSite As stdole.IUnknown)
    
    Dim oSiteOW As IOleWindow
    
    On Error Resume Next
    
    ' Store the new site object
    Set m_Site = pUnkSite
    
    If Not m_Site Is Nothing Then
        
        ' Create the band window
        Set m_Band = New frmPwdViewBand
        
        Set m_Band.IEWindow = FindIESite(m_Site)
        
        ' Get the IOleWindow interface of the band site
        Set oSiteOW = m_Site
        
        ' Move the window to the band site
        SetParent m_Band.hWnd, oSiteOW.GetWindow()
        
    Else
        
        ' Destroy the window
        Set m_Band = Nothing
        Set g_oApp = Nothing
    End If
    
End Sub

Private Function IOleWindow_GetWindow() As Long
   
   On Error Resume Next
   
   IOleWindow_GetWindow = m_Band.hWnd
   
End Function

Private Sub IPersist_GetClassID(pClassID As IOWS.UUID)
   
   On Error Resume Next
   
   ' Return the CLSID of this class
   CLSIDFromProgID "PwdViewBands.ExplBand", pClassID
 
End Sub

Private Sub IPersistStream_GetClassID(pClassID As IOWS.UUID)

   On Error Resume Next
   
   IPersist_GetClassID pClassID
   
End Sub

Private Sub IPersistStream_GetSizeMax(ByVal pcbSize As Long)

   Err.Raise E_NOTIMPL
   
End Sub

Private Sub IPersistStream_IsDirty()

   Err.Raise E_NOTIMPL
   
End Sub

Private Sub IPersistStream_Load(ByVal pStm As stg.IStream)
   
End Sub

Private Sub IPersistStream_Save(ByVal pStm As stg.IStream, ByVal fClearDirty As Long)
  
End Sub
