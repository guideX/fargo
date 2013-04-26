Attribute VB_Name = "mdlWindowChange"
Option Explicit
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const MF_BYPOSITION = &H400&
Private Const SC_CLOSE = &HF060&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
Private Const WS_CHILD = &H40000000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_CLIENTEDGE = 512
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOOLWINDOW = 128
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_maximize = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000
Private Const WS_POPUP = &H80000000
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

Public Sub ChangeWin(ByVal lhWnd As Long, Optional ByVal Border As Boolean = True, Optional ByVal TitleBar As Boolean = True, Optional ByVal Maximize As Boolean = True, Optional ByVal Minimize As Boolean = True, Optional ByVal SystemMenu As Boolean = True, Optional ByVal ThickFrame As Boolean = True, Optional ByVal VScroll As Boolean = False, Optional ByVal HScroll As Boolean = False, Optional ByVal ExTransparent As Boolean = False, Optional ByVal ExClientEdge As Boolean = False, Optional ByVal ExToolWindow As Boolean = False)
Dim l As Long
'Dim l As Long, lExStyle As Long
' lExStyle& = IIf(ExTransparent, WS_EX_TRANSPARENT, 0) Or IIf(ExTopMost, WS_EX_TOPMOST, 0)
Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
l& = GetWindowLong(lhWnd&, GWL_STYLE)
If ThickFrame = True Then l& = l& Or WS_THICKFRAME Else l& = l& And Not WS_THICKFRAME
If Border = True Then l& = l& Or WS_BORDER Else l& = l& And Not WS_BORDER
If TitleBar = True Then l& = l& Or WS_CAPTION Else l& = l& And Not WS_CAPTION
If Maximize = True Then l& = l& Or WS_MAXIMIZEBOX Else l& = l& And Not WS_MAXIMIZEBOX
If Minimize = True Then l& = l& Or WS_MINIMIZEBOX Else l& = l& And Not WS_MINIMIZEBOX
If SystemMenu = True Then l& = l& Or WS_SYSMENU Else l& = l& And Not WS_SYSMENU
If VScroll = True Then l& = l& Or WS_VSCROLL Else l& = l& And Not WS_VSCROLL
If HScroll = True Then l& = l& Or WS_HSCROLL Else l& = l& And Not WS_HSCROLL
Call SetWindowLong(lhWnd&, GWL_STYLE, l&)
Call SetWindowPos(lhWnd&, 0, 0, 0, 0, 0, swpFlags)
l& = 0
l& = GetWindowLong(lhWnd&, GWL_EXSTYLE)
If ExTransparent = True Then l& = l& Or WS_EX_TRANSPARENT Else l& = l& And Not WS_EX_TRANSPARENT
If ExClientEdge = True Then l& = l& Or WS_EX_CLIENTEDGE Else l& = l& And Not WS_EX_CLIENTEDGE
'If ExToolWindow = True Then l& = l& Or WS_EX_TOOLWINDOW Else l& = l& And Not WS_ToolWindow
Call SetWindowLong(lhWnd&, GWL_EXSTYLE, l&)
Call SetWindowPos(lhWnd&, 0, 0, 0, 0, 0, swpFlags)
End Sub

Private Function FlipProp(ByVal lhWnd As Long, ByVal lBit As Long, ByVal bValue As Boolean, Optional dStyle As Long = GWL_STYLE) As Boolean
Dim l As Long
Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
l& = GetWindowLong(lhWnd&, dStyle&)
If bValue = True Then l& = l& Or lBit& Else l& = l& And Not lBit&
Call SetWindowLong(lhWnd&, dStyle&, l&)
Call SetWindowPos(lhWnd&, 0, 0, 0, 0, 0, swpFlags)
FlipProp = l& = GetWindowLong(lhWnd&, dStyle&)
End Function

Public Sub ClearSysMenu(ByVal lhWnd As Long)
Dim l As Long, o As Long, n As Long
l = GetSystemMenu(lhWnd&, False)
Do
    DoEvents
    o& = GetMenuItemID(l, n&)
    If o& <> SC_CLOSE Then Call RemoveMenu(l, n&, MF_BYPOSITION): n& = n& - 1
    n& = n& + 1
    If n& >= GetMenuItemCount(l&) - 1 Then Exit Do
Loop
'Call RemoveMenu(l, 0, MF_BYPOSITION)
End Sub
