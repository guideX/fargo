VERSION 5.00
Begin VB.Form frmToolBox 
   BorderStyle     =   0  'None
   Caption         =   "Tools"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   870
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
   Icon            =   "frmToolBox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   58
   ShowInTaskbar   =   0   'False
   Begin Fargo.XPButton cmdMouse 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   16777215
      MPTR            =   0
      MICON           =   "frmToolBox.frx":000C
      PICN            =   "frmToolBox.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   -1  'True
   End
   Begin Fargo.ctlFormDragger FormDragger1 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      Top             =   0
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   503
   End
   Begin VB.PictureBox cmdLabel 
      Height          =   375
      Left            =   435
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      ToolTipText     =   "Label"
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdFrame 
      Height          =   375
      Left            =   45
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      ToolTipText     =   "Frame"
      Top             =   3930
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdPicture 
      Height          =   375
      Left            =   45
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      ToolTipText     =   "PictureBox"
      Top             =   3150
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdTextBox 
      Height          =   375
      Left            =   435
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      ToolTipText     =   "TextBox"
      Top             =   3150
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdCheckBox 
      Height          =   375
      Left            =   45
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      ToolTipText     =   "CheckBox"
      Top             =   3540
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox cmdOption 
      Height          =   375
      Left            =   435
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      ToolTipText     =   "OptionBox"
      Top             =   3540
      Visible         =   0   'False
      Width           =   375
   End
   Begin Fargo.XPButton cmdButton 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   16777215
      MPTR            =   0
      MICON           =   "frmToolBox.frx":0228
      PICN            =   "frmToolBox.frx":0244
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   -1  'True
   End
End
Attribute VB_Name = "frmToolBox"
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
Enum eCurrentToolBoxType
    oOther = 0
    oLabel = 1
    oPictureBox = 2
    oCommandButton = 3
    oOptionBox = 4
    oTextBox = 5
    oFrame = 6
    oCheckBox = 7
    oSkinnedButton = 8
End Enum
Private lCurrentToolBoxType As eCurrentToolBoxType

Public Sub DeselectToolBoxTools()
cmdButton.Value = False
cmdMouse.Value = True
End Sub

Public Function GetCurrentToolBoxType() As eCurrentToolBoxType
GetCurrentToolBoxType = lCurrentToolBoxType
End Function

Public Sub SetCurrentToolBoxType(lType As eCurrentToolBoxType)
lCurrentToolBoxType = lType
End Sub

Public Function ReturnDockedWidth() As Long
ReturnDockedWidth = lDockedWidth
End Function

Public Function ReturnDockedHeight() As Long
ReturnDockedHeight = lDockedHeight
End Function

Public Sub FormDrag(lFormName As Form)
ReleaseCapture
Call SendMessage(lFormName.hWnd, &HA1, 2, 0&)
End Sub

Private Sub StoreFormDimensions()
If Not bMoving Then
    If bDocked Then
        lDockedWidth = Me.Width
        lDockedHeight = Me.Height
    Else
        lFloatingLeft = Me.Left
        lFloatingTop = Me.Top
        lFloatingWidth = Me.Width
        lFloatingHeight = Me.Height
    End If
End If
End Sub

Private Sub Command1_Click()
Me.BorderStyle = 3
End Sub

Private Sub cmdButton_Click()
lCurrentToolBoxType = oCommandButton
mdiMain.ActiveForm.SetFocus
cmdMouse.Value = False
'mdiMain.ActiveForm.NewButton
End Sub

Private Sub cmdButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    mdiMain.ActiveForm.NewButton
End If
End Sub

Private Sub cmdCheckBox_Click()
mdiMain.ActiveForm.NewCheckBox
End Sub

Private Sub cmdCheck_Click()

End Sub

Private Sub cmdFrame_Click()
mdiMain.ActiveForm.NewFrame
End Sub

Private Sub cmdLabel_Click()
mdiMain.ActiveForm.NewLabel
End Sub

Private Sub cmdMouse_Click()
If cmdMouse.Value = False Then cmdMouse.Value = True
mdiMain.ActiveForm.SetFocus
cmdButton.Value = False
End Sub

Private Sub cmdOption_Click()
mdiMain.ActiveForm.NewOptionButton
End Sub

Private Sub cmdPicture_Click()
mdiMain.ActiveForm.NewPicture
End Sub

Private Sub cmdTextBox_Click()
mdiMain.ActiveForm.NewTextBox
End Sub

Public Sub EnableToolBox(lEnabled As Boolean)
Select Case lEnabled
Case True
    frmToolBox.cmdButton.Enabled = True
    cmdMouse.Enabled = True
    'frmToolBox.cmdCheckBox.Enabled = True
    ''frmToolBox.cmdFrame.Enabled = True
    ''frmToolBox.cmdLabel.Enabled = True
    'frmToolBox.cmdOption.Enabled = True
    'frmToolBox.cmdPicture.Enabled = True
    ''frmToolBox.cmdTextBox.Enabled = True
Case False
    frmToolBox.cmdButton.Enabled = False
    cmdMouse.Enabled = False
    'frmToolBox.cmdCheckBox.Enabled = False
    'frmToolBox.cmdFrame.Enabled = False
    'frmToolBox.cmdLabel.Enabled = False
    'frmToolBox.cmdOption.Enabled = False
    'frmToolBox.cmdPicture.Enabled = False
    'frmToolBox.cmdTextBox.Enabled = False
End Select
End Sub

Private Sub Form_Load()
Dim i As Integer
FormDragger1.Caption = "ToolBox"
Caption = "ToolBox"
lDockedWidth = mdiMain.picToolBox.ScaleWidth + (8 * Screen.TwipsPerPixelX)
lDockedHeight = mdiMain.picToolBox.ScaleHeight + (8 * Screen.TwipsPerPixelY)
lFloatingLeft = Me.Left
lFloatingTop = Me.Top
lFloatingWidth = Me.Width
lFloatingHeight = Me.Height
bDocked = True
DockToolBox
End Sub

Public Sub UnDockToolBox()
mdiMain.picToolBox.Visible = False
SetParent hWnd, 0
End Sub

Public Sub DockToolBox()
mdiMain.picToolBox.Visible = True
SetParent hWnd, mdiMain!picToolBox.hWnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
mdiMain.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SetWindowWord(Me.hWnd, SWW_HPARENT, 0&)
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then StoreFormDimensions
mdiMain.picToolBox.Width = Me.Width
End Sub

Private Sub FormDragger1_DblClick()
bMoving = True
If bDocked = True Then
    Me.Visible = False
    bDocked = False
    SetParent Me.hWnd, 0
    Me.Move lFloatingLeft, lFloatingTop, lFloatingWidth, lFloatingHeight
    mdiMain!picToolBox.Visible = False
    Me.Visible = True
    Call SetWindowWord(Me.hWnd, SWW_HPARENT, mdiMain.hWnd)
Else
    bDocked = True
    SetParent Me.hWnd, mdiMain!picToolBox.hWnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    mdiMain!picToolBox.Visible = True
End If
bMoving = False
mdiMain.ActivateResize
End Sub

Private Sub FormDragger1_FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
Dim rct As RECT
GetWindowRect mdiMain!picToolBox.hWnd, rct
With rct
    .Left = .Left - 4
    .Top = .Top - 4
    .Right = .Right + 4
    .Bottom = .Bottom + 4
End With
If PtInRect(rct, FormLeft, FormTop) Then
    bDocked = True
    SetParent Me.hWnd, mdiMain!picToolBox.hWnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    mdiMain!picToolBox.Visible = True
Else
    Me.Visible = False
    bDocked = False
    SetParent Me.hWnd, 0
    Me.Move FormLeft * Screen.TwipsPerPixelX, FormTop * Screen.TwipsPerPixelY, lFloatingWidth, lFloatingHeight
    mdiMain!picToolBox.Visible = False
    Me.Visible = True
    Call SetWindowWord(Me.hWnd, SWW_HPARENT, mdiMain.hWnd)
End If
bMoving = False
StoreFormDimensions
mdiMain.ActivateResize
End Sub

Private Sub FormDragger1_FormMoved(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
Dim rct As RECT
bMoving = True
GetWindowRect mdiMain!picToolBox.hWnd, rct
With rct
    .Left = .Left - 4
    .Top = .Top - 4
    .Right = .Right + 4
    .Bottom = .Bottom + 4
End With
If PtInRect(rct, FormLeft, FormTop) Then
    FormWidth = lDockedWidth / Screen.TwipsPerPixelX
    FormHeight = lDockedHeight / Screen.TwipsPerPixelY
Else
    FormWidth = lFloatingWidth / Screen.TwipsPerPixelX
    FormHeight = lFloatingHeight / Screen.TwipsPerPixelY
End If
End Sub

Private Sub FormDragger1_mGotFocus()
mdiMain.SetFocus
End Sub

Private Sub FormDragger1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdiMain.SetFocus
End Sub

Private Sub OsenXPButton2_Click()

End Sub

Private Sub OsenXPButton1_Click()

End Sub

Private Sub txtTextBox_Click()

End Sub

Private Sub txtText_Click()

End Sub
