Attribute VB_Name = "ContextMenu"
Public Declare Function InsertMenu Lib "user32" _
               Alias "InsertMenuA" (ByVal HMENU As Long, _
               ByVal nPosition As Long, ByVal wFlags As Long, _
               ByVal wIDNewItem As Long, _
               ByVal lpNewItem As String) As Long
               
 'Menu Constants
  Public Const MF_BYPOSITION = &H400&
  Public Const MF_STRING = &H0&
  Public Const MF_SEPARATOR = &H800&
               
Public Function QueryContextMenuX( _
                ByVal This As IContextMenu, _
                ByVal HMENU As Long, _
                ByVal indexMenu As Long, _
                ByVal idCmdFirst As Long, _
                ByVal idCmdLast As Long, _
                ByVal uFlags As Long) As Long

Dim sMenuItem As String
Dim idCmd As Long

idCmd = idCmdFirst

sMenuItem = "&Open"
Call InsertMenu(HMENU, indexMenu, MF_STRING Or MF_BYPOSITION, _
                idCmd, sMenuItem)
idCmd = idCmd + 1
indexMenu = indexMenu + 1

sMenuItem = "&Play"
Call InsertMenu(HMENU, indexMenu, MF_STRING Or MF_BYPOSITION, _
                idCmd, sMenuItem)
idCmd = idCmd + 1
indexMenu = indexMenu + 1

sMenuItem = "&Stop"
Call InsertMenu(HMENU, indexMenu, MF_STRING Or MF_BYPOSITION, _
                idCmd, sMenuItem)
idCmd = idCmd + 1
indexMenu = indexMenu + 1


QueryContextMenuX = indexMenu



End Function

