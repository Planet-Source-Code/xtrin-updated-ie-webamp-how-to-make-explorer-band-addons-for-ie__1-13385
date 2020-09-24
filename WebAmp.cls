VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebAmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_pSite As IUnknownVB
Private m_ContainerWnd As Long
Private m_bandWnd As Long
Private m_pOldQueryContextMenu As Long

Private Const IID_IWebBrowserApp = "{0002DF05-0000-0000-C000-000000000046}"
Private Const IID_IWebBrowser2 = "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"

Implements IDeskBand
Implements IInputObject
Implements IObjectWithSite
Implements IContextMenu


Private Sub Class_Initialize()
    
    Dim pContextMenu As IContextMenu
    Set pContextMenu = Me
    
    m_pOldQueryContextMenu = SwapVtableEntry(ObjPtr(pContextMenu), 4, AddressOf QueryContextMenuX)
    
End Sub

Private Sub Class_Terminate()

    Dim pContextMenu As IContextMenu
    Set pContextMenu = Me
    
    m_pOldQueryContextMenu = SwapVtableEntry(ObjPtr(pContextMenu), 4, m_pOldQueryContextMenu)
    
End Sub


Private Sub IContextMenu_GetCommandString(ByVal idCmd As VBShellLib.UINT, ByVal uType As VBShellLib.UINT, ByVal pwReserved As VBShellLib.UINT, ByVal pszName As VBShellLib.LPSTRVB, ByVal cchMax As VBShellLib.UINT)
    Err.Raise E_NOTIMPL
End Sub


Private Sub IContextMenu_InvokeCommand(ByVal lpcmi As VBShellLib.LPCMINVOKECOMMANDINFO)
        
    'Let the band handle the menu implementation
    frmBand.MenuHandler lpcmi

End Sub


Private Sub IContextMenu_QueryContextMenu(ByVal HMENU As VBShellLib.HMENU, ByVal indexMenu As VBShellLib.UINT, ByVal idCmdFirst As VBShellLib.UINT, ByVal idCmdLast As VBShellLib.UINT, ByVal uFlags As VBShellLib.QUERYCONTEXTMENUFLAGS)
    'Swapped
End Sub

Private Sub IDeskBand_CloseDW(ByVal dwReserved As Long)
    'The band form will not trigger it's unload or terminate event
    'unless we shut it down here
    
    Unload frmBand
    
End Sub

Private Sub IDeskBand_ContextSensitiveHelp(ByVal fEnterMode As Boolean)
    Err.Raise E_NOTIMPL
End Sub

Private Sub IDeskBand_GetBandInfo(ByVal dwBandID As Long, ByVal dwViewMode As Long, ByVal pdbi As VBShellLib.LPDESKBANDINFO)

    Dim dbi As DESKBANDINFO
    
    If pdbi = 0 Then
        Exit Sub
    End If
    
    CopyMemory dbi, ByVal pdbi, Len(dbi)
    
    If (dbi.dwMask And DBIM_MINSIZE) Then
        dbi.ptMinSize.x = 10&
        dbi.ptMinSize.y = 50&
    End If

    If (dbi.dwMask And DBIM_MAXSIZE) Then
        dbi.ptMaxSize.x = -1&
        dbi.ptMaxSize.y = -1&
    End If

    If (dbi.dwMask And DBIM_INTEGRAL) Then
        dbi.ptIntegral.x = 1&
        dbi.ptIntegral.y = 1&
    End If

    If (dbi.dwMask And DBIM_ACTUAL) Then
        dbi.ptActual.x = 0&
        dbi.ptActual.y = 0&
    End If

    If (dbi.dwMask And DBIM_TITLE) Then
        Dim title() As Byte
        title = "IE WebAmp" & vbNullChar
        CopyMemory dbi.wszTitle(0), title(0), UBound(title) + 1
    End If

    If (dbi.dwMask And DBIM_MODEFLAGS) Then
        dbi.dwModeFlags = DBIMF_VARIABLEHEIGHT
    End If

    If (dbi.dwMask And DBIM_BKCOLOR) Then
        'Use the default background color by removing
        'DBIM_BKCOLOR flag and setting crBkgnd
    End If
    
    CopyMemory ByVal pdbi, dbi, Len(dbi)

End Sub

Private Function IDeskBand_GetWindow() As Long
    IDeskBand_GetWindow = m_bandWnd
End Function

Private Sub IDeskBand_ResizeBorderDW(ByVal prcBorder As Long, ByVal punkToolbarSite As Long, ByVal fReserved As Boolean)
    Err.Raise E_NOTIMPL
End Sub

Private Sub IDeskBand_ShowDW(ByVal fShow As Boolean)
    If (fShow) Then
        ShowWindow m_bandWnd, SW_SHOW
    Else
        ShowWindow m_bandWnd, SW_HIDE
    End If
End Sub

Private Sub IInputObject_HasFocusIO()
    Err.Raise E_NOTIMPL
End Sub

Private Sub IInputObject_TranslateAcceleratorIO(ByVal lpMsg As lpMsg)
    Err.Raise E_NOTIMPL
End Sub

Private Sub IInputObject_UIActivateIO(ByVal fActivate As Boolean, ByVal lpMsg As lpMsg)
    If (fActivate) Then
        SetFocus m_bandWnd
    End If
End Sub

Private Sub IObjectWithSite_GetSite(ByVal priid As LPGUID, ppvSite As LPVOID)
    m_pSite.QueryInterface priid, ppvSite
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pUnkSite As IUnknownVB)

    Dim isp As IServiceProvider
    Dim oleWnd As IOleWindow
        
    Dim wba As GUID 'IWebBrowserApp
    Dim wb2 As GUID 'IWebBrowser2
        
    Dim dwStyle As Long
    
    If Not (pUnkSite Is Nothing) Then
    
        If Not (m_pSite Is Nothing) Then
            Set m_pSite = Nothing
        End If
        
        Set m_pSite = pUnkSite
        Set oleWnd = pUnkSite
        
        'QueryInterface for IServiceProvider
        Set isp = pUnkSite
        
        'Query service provider to get IWebBrowser2 (InternetExplorer)
        CLSIDFromString StrPtr(IID_IWebBrowserApp), wba
        CLSIDFromString StrPtr(IID_IWebBrowser2), wb2
        
        Load frmBand
        'Set frmBand.InternetExplorer = isp.QueryService(VarPtr(wba), VarPtr(wb2))
        
        Set isp = Nothing
        
        If Not (oleWnd Is Nothing) Then
            
            m_ContainerWnd = oleWnd.GetWindow
            m_bandWnd = frmBand.hwnd
            
            dwStyle = GetWindowLong(m_bandWnd, GWL_STYLE)
            dwStyle = dwStyle Or WS_CHILD Or WS_CLIPSIBLINGS
            SetWindowLong m_bandWnd, GWL_STYLE, dwStyle
            SetParent m_bandWnd, m_ContainerWnd
              
        End If
        
        Set oleWnd = Nothing
    Else
        Set m_pSite = Nothing
    End If
   
End Sub

