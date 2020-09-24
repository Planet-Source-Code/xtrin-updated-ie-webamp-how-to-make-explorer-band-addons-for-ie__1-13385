Attribute VB_Name = "BandObject"
Option Explicit

Public Declare Function CLSIDFromProgID Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As GUID) As Long
Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As GUID) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long


'ShowWindow Constants
Public Const SW_SHOW = 5
Public Const SW_HIDE = 0

'Window Style Constants
Public Const GWL_STYLE = (-16)
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPSIBLINGS = &H4000000

'Shell Constants
Public Const PAGE_EXECUTE_READWRITE = &H40&


Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Type DESKBANDINFO
    dwMask As Long
    ptMinSize As POINTAPI
    ptMaxSize As POINTAPI
    ptIntegral As POINTAPI
    ptActual As POINTAPI
    wszTitle(511) As Byte
    dwModeFlags As Long
    crBkgnd As Long
End Type

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
          
Public Type CMINVOKECOMMANDINFO
    cbSize As Long          ' sizeof(CMINVOKECOMMANDINFO)
    fMask As Long           ' any combination of CMIC_MASK_*
    hwnd  As Long           ' might be NULL (indicating no owner window)
    lpVerb As Long          ' either a string or MAKEINTRESOURCE(idOffset)
    lpParameters As Long    ' might be NULL (indicating no parameter)
    lpDirectory As Long     ' might be NULL (indicating no specific directory)
    nShow As Long           ' one of SW_ values for ShowWindow() API
    dwHotKey As Long
    hIcon As Long
End Type

Public Const E_NOTIMPL = &H80004001

'Functions
Public Function SwapVtableEntry(pObj As Long, EntryNum As Integer, ByVal lpfn As Long) As Long
    Dim lOldAddr As Long
    Dim lpVtable As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long
    
    CopyMemory lpVtable, ByVal pObj, 4
    lpfnAddr = lpVtable + (EntryNum - 1) * 4
    CopyMemory lOldAddr, ByVal lpfnAddr, 4
    
    Call VirtualProtect(lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
    CopyMemory ByVal lpfnAddr, lpfn, 4
    Call VirtualProtect(lpfnAddr, 4, lOldProtect, lOldProtect)
    
    SwapVtableEntry = lOldAddr
End Function


