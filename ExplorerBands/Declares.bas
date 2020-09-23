Attribute VB_Name = "modDeclares"
'------------------------------------------------------------------------------
' PwdViewBands
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

Public g_oApp As PasswordViewBus.Application


Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal Ln As Long)

Public Const SW_HIDE = 0
Public Const SW_SHOW = 1

Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Const GWL_STYLE = (-16)

Public Const WS_CHILD = &H40000000

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_BTNFACE = 15


'
' FindIESite
'
' Returns the explorer window that contains
' the band site
'
' Parameters:
'
' BandSite    IOleWindow interface of the band site
'
Public Function FindIESite(ByVal BandSite As IUnknown) As IWebBrowserApp
    
    Dim oServiceProvider As IServiceProvider
    Dim oShellBrowser As IUnknown
    Dim IID_IWebBrowserApp As UUID
    Dim SID_SInternetExplorer As UUID
    
    ' Convert IID and SID
    ' from strings to UUID UDTs
    CLSIDFromString sIID_IWebBrowserApp, IID_IWebBrowserApp
    CLSIDFromString sSID_SInternetExplorer, SID_SInternetExplorer
    
    ' Get IServiceProvider interface
    ' of the band site
    Set oServiceProvider = BandSite
      
    ' Get the InternetExplorer
    ' object through IServiceProvider
    Set FindIESite = oServiceProvider.QueryService(SID_SInternetExplorer, IID_IWebBrowserApp)
         
End Function
