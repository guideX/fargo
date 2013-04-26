VERSION 5.00
Begin VB.Form frmProperties 
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   2565
   ClientWidth     =   2670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
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
   ScaleHeight     =   2175
   ScaleWidth      =   2670
   Begin VisualCodeDesigner.ctlFormDragger ctlFormDragger1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   688
      Caption         =   "Properties"
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal lLeft As Long, ByVal lTop As Long) As Long
Private Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Private Const SWW_HPARENT = -8
Private Const HTRIGHT = 11
Private bMoving As Boolean
Private lFloatingLeft As Long
Private lFloatingTop As Long
Private lFloatingWidth As Long
Private lFloatingHeight As Long
Private bDocked As Boolean
Private lDockedWidth As Long
Private lDockedHeight As Long

Private Sub Command1_Click()
ctlPropertyList1.ClearPropList
End Sub

Private Sub ctlFormDragger1_DblClick()
bMoving = True
If bDocked = True Then
    Me.Visible = False
    bDocked = False
    SetParent Me.hWnd, 0
    Me.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
    mdiMain!picProporties.Visible = False
    Me.Visible = True
    Call SetWindowWord(Me.hWnd, SWW_HPARENT, mdiMain.hWnd)
Else
    bDocked = True
    SetParent Me.hWnd, mdiMain!picProporties.hWnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    mdiMain!picProporties.Visible = True
End If
bMoving = False
mdiMain.ActivateResize
End Sub

Private Sub Form_Load()
Dim msg As String, i As Integer, msg2 As String
msg = "0 - You suck" & vbCrLf & "1 - You blow" & vbCrLf & "2 - You suck hard"
msg2 = "fuck" & vbCrLf & "shit" & vbCrLf & "sex" & vbCrLf & "teenz"
For i = 0 To 30
    ctlPropertyList1.AddProporty Trim(Str(i)), "Test", nDropDown, msg
Next i
Form_Resize

lDockedWidth = mdiMain.picProporties.ScaleWidth + (8 * Screen.TwipsPerPixelX)
lDockedHeight = mdiMain.picProporties.ScaleHeight + (8 * Screen.TwipsPerPixelY)
lFloatingLeft = Me.Left
lFloatingTop = Me.Top
lFloatingWidth = Me.Width
lFloatingHeight = Me.Height
bDocked = True
SetParent hWnd, mdiMain!picProporties.hWnd

End Sub

Private Sub Form_Resize()
ctlPropertyList1.ResizeControl Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
ctlPropertyList1.Width = Me.ScaleWidth
ctlPropertyList1.Height = Me.ScaleHeight
End Sub

