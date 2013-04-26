VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{004BB5D5-7D55-4298-B546-5ABB5C35F3AC}#1.0#0"; "NexgenTab.ocx"
Begin VB.Form frmProporties 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   11280
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   2595
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
   Icon            =   "frmProporties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProporties.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProporties.frx":57FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProporties.frx":AFF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProporties.frx":107E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProporties.frx":15FD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin nTab.nTabControl nTabControl1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5106
      TabCount        =   2
      TabCaption(0)   =   "Properties"
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "ctlPropertyList1"
      TabCaption(1)   =   "Project"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "tvwProject"
      TabTheme        =   2
      InActiveTabBackStartColor=   -2147483626
      InActiveTabBackEndColor=   -2147483626
      InActiveTabForeColor=   -2147483631
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   -2147483628
      TabStripBackColor=   -2147483626
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      Begin Fargo.ctlPropertyList ctlPropertyList1 
         Height          =   2415
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   4260
      End
      Begin MSComctlLib.TreeView tvwProject 
         Height          =   1815
         Left            =   -75000
         TabIndex        =   2
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3201
         _Version        =   393217
         Indentation     =   176
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   0
      End
   End
   Begin VB.Timer tmrSetTab 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin Fargo.ctlFormDragger FormDragger1 
      Align           =   1  'Align Top
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   529
   End
End
Attribute VB_Name = "frmProporties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Enum ePropertyType
    oOther = 0
    oLabel = 1
    oPictureBox = 2
    oCommandButton = 3
    oOptionBox = 4
    oTextBox = 5
    oFrame = 6
    oCheckBox = 7
    oSkinnedButton = 8
    oForm = 9
End Enum
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
Private lWindIndex As Integer
Private lPropertyIndex As Integer
Private lPropertyType As ePropertyType
Private Type gPropertiesList
    pName As String
    pValue As String
End Type
Private lPropertiesList As gPropertiesList

Public Sub SetObjectIndex(lIndex As Integer)
lPropertyIndex = lIndex
End Sub

Public Sub SetObjectType(lObj As eObjectType)
lPropertyType = lObj
End Sub

Public Function ReturnDockedWidth() As Long
ReturnDockedWidth = lDockedWidth
Exit Function
End Function

Public Function ReturnDockedHeight() As Long
ReturnDockedHeight = lDockedHeight
End Function

Private Sub FormDrag(lFormName As Form)
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

Public Sub SetWindIndex(lIndex As Integer)
lWindIndex = lIndex
End Sub

Private Sub ctlPropertyList1_KeyPress(lKey As Integer)
Dim msg As String
'Stop
If lKey = 13 Then
    If Len(lPropertiesList.pName) <> 0 Then
'        MsgBox lPropertyType
        Select Case lPropertyType
        Case 0
            'Stop
            Select Case Trim(LCase(lPropertiesList.pName))
            Case "(name)"
                mdiMain.SetFormName lWindIndex, lPropertiesList.pValue
            Case "caption"
                mdiMain.SetFormCaption lWindIndex, lPropertiesList.pValue
            Case "backcolor"
                mdiMain.SetFormBackColor lWindIndex, lPropertiesList.pValue
            End Select
            
        Case oCommandButton
            Select Case Trim(LCase(lPropertiesList.pName))
            Case "(name)"
                mdiMain.SetButtonName lWindIndex, lPropertyIndex, lPropertiesList.pValue
            Case "caption"
                mdiMain.SetButtonCaption lWindIndex, lPropertyIndex, lPropertiesList.pValue
            Case "left"
                
                mdiMain.SetButtonLeft lWindIndex, lPropertyIndex, Int(lPropertiesList.pValue)
                
                'Stop
                'mdiMain.ActiveForm.SetHashMarksOnCurrentObject
            Case "top"
                mdiMain.SetButtonTop lWindIndex, lPropertyIndex, Int(lPropertiesList.pValue)
            Case "width"
                mdiMain.SetButtonWidth lWindIndex, lPropertyIndex, Int(lPropertiesList.pValue)
            Case "height"
                mdiMain.SetButtonHeight lWindIndex, lPropertyIndex, Int(lPropertiesList.pValue)
            End Select
        Case oForm
        '    MsgBox "form"
        End Select
    End If
End If
End Sub

Private Sub ctlPropertyList1_ValueChanged(lIndex As Integer, lName As String, lValue As String)
lPropertiesList.pName = lName
lPropertiesList.pValue = lValue
End Sub

Private Sub Form_Load()
Dim i As Integer
FormDragger1.Caption = "Properties"
Form_Resize
tvwProject.Visible = True
lDockedWidth = mdiMain.picProporties.ScaleWidth + (8 * Screen.TwipsPerPixelX)
lDockedHeight = mdiMain.picProporties.ScaleHeight + (8 * Screen.TwipsPerPixelY)
lFloatingLeft = Me.Left
lFloatingTop = Me.Top
lFloatingWidth = Me.Width
lFloatingHeight = Me.Height
bDocked = True
SetParent hWnd, mdiMain!picProporties.hWnd
tmrSetTab.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
mdiMain.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call SetWindowWord(Me.hWnd, SWW_HPARENT, 0&)
End Sub

Private Sub Form_Resize()
nTabControl1.Width = Me.ScaleWidth
nTabControl1.Height = Me.ScaleHeight - 550
If Me.WindowState <> vbMinimized Then StoreFormDimensions
If ctlPropertyList1.Left <> 40 Then ctlPropertyList1.Left = 40
If ctlPropertyList1.Top <> 360 Then ctlPropertyList1.Top = 360
ctlPropertyList1.Width = Me.ScaleWidth - 100
ctlPropertyList1.Height = Me.ScaleHeight - 500
ctlPropertyList1.ResizeControl 40, 360, Me.ScaleWidth - 100, Me.ScaleHeight - 920
tvwProject.Width = Me.ScaleWidth - 60
tvwProject.Height = Me.ScaleHeight - 940

'ctlPropertyList1.Height = Me.ScaleHeight
'ctlPropertyList1.Width = Me.ScaleWidth
'ctlPropertyList1.ResizeControl Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
'ctlPropertyList1.AddProporty "Fuck", "Shit", nDropDown, "Fuck" & vbCrLf & "Shit" & vbCrLf & "DogShit"
End Sub

Private Sub FormDragger1_DblClick()
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

Private Sub FormDragger1_FormDropped(FormLeft As Long, FormTop As Long, FormWidth As Long, FormHeight As Long)
Dim rct As RECT
GetWindowRect mdiMain!picProporties.hWnd, rct
With rct
    .Left = .Left - 4
    .Top = .Top - 4
    .Right = .Right + 4
    .Bottom = .Bottom + 4
End With
If PtInRect(rct, FormLeft, FormTop) Then
    bDocked = True
    SetParent Me.hWnd, mdiMain!picProporties.hWnd
    Me.Move -4 * Screen.TwipsPerPixelX, -4 * Screen.TwipsPerPixelY, lDockedWidth, lDockedHeight
    mdiMain!picProporties.Visible = True
Else
    Me.Visible = False
    bDocked = False
    SetParent Me.hWnd, 0
    Me.Move FormLeft * Screen.TwipsPerPixelX, FormTop * Screen.TwipsPerPixelY, lFloatingWidth, lFloatingHeight
    mdiMain!picProporties.Visible = False
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
GetWindowRect mdiMain!picProporties.hWnd, rct
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

Private Sub Timer1_Timer()
'FormDragger1.Caption = Me.Left & " - " & Me.Top
FormDragger1.Caption = Me.Height
End Sub

Private Sub nTabControl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mdiMain.SetFocus
End Sub

Private Sub tmrSetTab_Timer()
nTabControl1.ActiveTab = 0
tmrSetTab.Enabled = False
End Sub

Private Sub tvwProject_DblClick()
If Len(tvwProject.SelectedItem.Tag) <> 0 Then
    Select Case LCase(Trim(tvwProject.SelectedItem.Parent.Text))
    Case "modules"
        mdiMain.SetFocusModule Int(Trim(tvwProject.SelectedItem.Tag)), True
    Case "forms"
        mdiMain.SetFocusForm Int(Trim(tvwProject.SelectedItem.Tag))
    End Select
End If
End Sub

Private Sub XTab1_TabSwitch(ByVal iLastActiveTab As Integer)
On Local Error Resume Next
If iLastActiveTab = 0 Then
    ctlPropertyList1.Visible = True
    tvwProject.Visible = False
    ctlPropertyList1.SetFocus
Else
    ctlPropertyList1.Visible = False
    tvwProject.Visible = True
    tvwProject.SetFocus
End If
If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub XTab1_Click()

End Sub
