VERSION 5.00
Begin VB.UserControl ctlFormDragger 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   CanGetFocus     =   0   'False
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   675
   ScaleWidth      =   720
End
Attribute VB_Name = "ctlFormDragger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Const BDR_SUNKENINNER = &H8
Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_RAISED = &H5
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Const VK_LBUTTON = &H1
Private Const PS_SOLID = 0
Private Const R2_NOTXORPEN = 10
Private Const BLACK_PEN = 7
Private Const SM_CYCAPTION = 4
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
Event FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
Event mGotFocus()
Const m_def_RepositionForm = True
Const m_def_Caption = ""
Dim m_RepositionForm As Boolean
Dim m_Caption As String

Private Sub UserControl_GotFocus()
RaiseEvent mGotFocus
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim l As Long, p As POINTAPI, o As Long
UserControl_Paint
o = UserControl.Extender.Parent.hWnd
If Button = vbLeftButton And X >= 0 And X <= UserControl.ScaleWidth And Y >= 0 And Y <= UserControl.ScaleHeight Then
    ReleaseCapture
    DragObject o
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub DragObject(ByVal hWnd As Long)
Dim p As POINTAPI, p2 As POINTAPI, r As RECT, r2 As RECT, l As Long, o As Long, n As Long, g As Long, l2 As Long, l3 As Long, b As Boolean
ReleaseCapture
GetWindowRect hWnd, r
n = r.Right - r.Left
g = r.Bottom - r.Top
GetCursorPos p
p2.X = p.X
p2.Y = p.Y
l2 = p.X - r.Left
l3 = p.Y - r.Top
With r2
    .Left = p.X - l2
    .Top = p.Y - l3
    .Right = .Left + n
    .Bottom = .Top + g
End With
o = 3
DrawDragRectangle r2.Left, r2.Top, r2.Right, r2.Bottom, o
Do While GetKeyState(VK_LBUTTON) < 0
    GetCursorPos p
    If p.X <> p2.X Or p.Y <> p2.Y Then
        p2.X = p.X
        p2.Y = p.Y
        DrawDragRectangle r2.Left, r2.Top, r2.Right, r2.Bottom, o
        RaiseEvent FormMoved(p.X - l2, p.Y - l3, n, g)
        With r2
            .Left = p.X - l2
            .Top = p.Y - l3
            .Right = .Left + n
            .Bottom = .Top + g
        End With
        DrawDragRectangle r2.Left, r2.Top, r2.Right, r2.Bottom, o
        b = True
    End If
    DoEvents
Loop
DrawDragRectangle r2.Left, r2.Top, r2.Right, r2.Bottom, o
If b Then
    If m_RepositionForm Then
        MoveWindow hWnd, r2.Left, r2.Top, r2.Right - r2.Left, r2.Bottom - r2.Top, True
    End If
    RaiseEvent FormDropped(r2.Left, r2.Top, r2.Right - r2.Left, r2.Bottom - r2.Top)
End If
End Sub

Private Sub DrawDragRectangle(ByVal X As Long, ByVal Y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long)
Dim l As Long, o As Long
o = CreatePen(PS_SOLID, lWidth, &HE0E0E0)
l = GetDC(0)
Call SelectObject(l, o)
Call SetROP2(l, R2_NOTXORPEN)
Call Rectangle(l, X, Y, X1, Y1)
Call SelectObject(l, GetStockObject(BLACK_PEN))
Call DeleteObject(o)
Call SelectObject(l, o)
Call ReleaseDC(0, l)
End Sub

Private Sub UserControl_InitProperties()
m_Caption = m_def_Caption
m_Caption = m_def_Caption
m_RepositionForm = m_def_RepositionForm
End Sub

Private Sub UserControl_Paint()
Dim l As Long
Dim msg As String
With UserControl
    .Cls
    .Extender.Align = vbAlignTop
    .Extender.Top = 0
    .Height = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
    If GetActiveWindow = UserControl.Extender.Parent.hWnd Then
        .ForeColor = vbTitleBarText
        l = vbActiveTitleBar
    Else
        .ForeColor = vbInactiveTitleBarText
        l = vbInactiveTitleBar
    End If
    UserControl.Line (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - (2 * Screen.TwipsPerPixelX), UserControl.ScaleHeight - Screen.TwipsPerPixelY), l, BF
    .CurrentX = 4 * Screen.TwipsPerPixelX
    .CurrentY = 3 * Screen.TwipsPerPixelY
    .Font.Name = "Tahoma"
    .Font.Bold = True
    msg = m_Caption
    If UserControl.TextWidth(msg) > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) Then
        Do While UserControl.TextWidth(msg & "...") > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) And Len(msg) > 0
            msg = Trim$(Left$(msg, Len(msg) - 1))
        Loop
        msg = msg & "..."
    End If
    UserControl.Print msg;
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
m_RepositionForm = PropBag.ReadProperty("RepositionForm", m_def_RepositionForm)
UserControl_Paint
End Sub

Private Sub UserControl_Resize()
UserControl_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
Call PropBag.WriteProperty("RepositionForm", m_RepositionForm, m_def_RepositionForm)
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/Returns the caption of the control."
Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
m_Caption = New_Caption
PropertyChanged "Caption"
UserControl_Paint
End Property

Public Property Get RepositionForm() As Boolean
Attribute RepositionForm.VB_Description = "Specifies whether the control should move the form to it's new location."
RepositionForm = m_RepositionForm
End Property

Public Property Let RepositionForm(ByVal New_RepositionForm As Boolean)
m_RepositionForm = New_RepositionForm
PropertyChanged "RepositionForm"
End Property

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
UserControl.Refresh
End Sub
