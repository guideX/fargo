VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Fargo"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   870
   ClientWidth     =   9855
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stsBottom 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   5025
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picObjectDesigner 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9855
      TabIndex        =   4
      Top             =   345
      Visible         =   0   'False
      Width           =   9855
      Begin Fargo.XPButton cmdCreateObject 
         Height          =   255
         Left            =   6240
         TabIndex        =   13
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Create"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "mdiMain.frx":0CCA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   7440
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.ComboBox cboObjectType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "mdiMain.frx":0CE6
         Left            =   720
         List            =   "mdiMain.frx":0CFF
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   60
         Width           =   1215
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   60
         Width           =   1335
      End
      Begin VB.TextBox txtCaption 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Object:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   90
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Caption:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   90
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4320
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0D43
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1067
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":138B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":16AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":19D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1CF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":201B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":233F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2663
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2987
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2CAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2FCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":32F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3617
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":393B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3C5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3F83
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picProporties 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   7155
      ScaleHeight     =   4305
      ScaleWidth      =   2700
      TabIndex        =   2
      Top             =   720
      Width           =   2700
   End
   Begin MSComctlLib.Toolbar tlbTop 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   13
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   8
            Style           =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   17
         EndProperty
      EndProperty
      Begin VB.ComboBox cboObject 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   4935
      End
   End
   Begin VB.PictureBox picToolBox 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   0
      ScaleHeight     =   4305
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Begin VB.Menu mnuForm 
            Caption         =   "Form"
         End
         Begin VB.Menu mnuNewProject 
            Caption         =   "Project"
         End
         Begin VB.Menu mnuNewModule 
            Caption         =   "Module"
         End
      End
      Begin VB.Menu mnuSep327089742 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenProject 
         Caption         =   "Open Project"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveProject 
         Caption         =   "Save Project As ..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Debug"
      Begin VB.Menu mnuStart 
         Caption         =   "Run"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "Compile"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToggleToolBox 
         Caption         =   "Toggle ToolBox"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToggleProporties 
         Caption         =   "Toggle Proporties"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep37463 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSnap 
         Caption         =   "Snap"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSnapTo 
         Caption         =   "Snap To"
         Visible         =   0   'False
         Begin VB.Menu mnuGridLine 
            Caption         =   "Gridlines"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuDock 
         Caption         =   "Dock"
         Begin VB.Menu mnuDockToolBox 
            Caption         =   "ToolBox"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuDockProperties 
            Caption         =   "Properties"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "Project"
      Visible         =   0   'False
      Begin VB.Menu mnuProjectProporties 
         Caption         =   "Customize"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Window"
      Begin VB.Menu mnuAutoMaximize 
         Caption         =   "Auto Maximize"
      End
      Begin VB.Menu mnuSep3976798263 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTileHorrizontal 
         Caption         =   "Tile Horrizontal"
      End
      Begin VB.Menu mnuTileVerticle 
         Caption         =   "Tile Verticle"
      End
      Begin VB.Menu mnuArrangeIcons 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuObjectMenu 
      Caption         =   "<FormObjectMenu>"
      Visible         =   0   'False
      Begin VB.Menu mnuCut1 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopy1 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste1 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuSep976329263 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuFormMenu 
      Caption         =   "<FormMenu>"
      Visible         =   0   'False
      Begin VB.Menu mnuViewCode 
         Caption         =   "View Code"
      End
      Begin VB.Menu mnuMenuEditor 
         Caption         =   "Menu Editor"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private lInterpreterLocation As String
Private lProject As clsProject
Private Enum eObjectType
    oOther = 0
    oLabel = 1
    oPictureBox = 2
    oCommandButton = 3
    oOptionBox = 4
    oTextBox = 5
    oFrame = 6
    oCheckBox = 7
End Enum

Public Sub SetFormBackColor(lIndex As Integer, lBackColor As String)
lProject.SetFormBackColor lIndex, lBackColor
End Sub

Public Sub SetFormCaption(lIndex As Integer, lName As String)
lProject.SetFormCaption lIndex, lName
End Sub

Public Sub SetFormName(lIndex As Integer, lName As String)
lProject.SetFormName lIndex, lName
End Sub

Public Sub SetButtonLeft(lFormIndex As Integer, lButtonIndex As Integer, lLeft As Integer)
lProject.SetButtonLeft lFormIndex, lButtonIndex, lLeft
End Sub

Public Sub SetButtonTop(lFormIndex As Integer, lButtonIndex As Integer, lTop As Integer)
lProject.SetButtonTop lFormIndex, lButtonIndex, lTop
End Sub

Public Sub SetButtonWidth(lFormIndex As Integer, lButtonIndex As Integer, lWidth As Integer)
lProject.SetButtonWidth lFormIndex, lButtonIndex, lWidth
End Sub

Public Sub SetButtonHeight(lFormIndex As Integer, lButtonIndex As Integer, lHeight As Integer)
lProject.SetButtonHeight lFormIndex, lButtonIndex, lHeight
End Sub

Public Sub ReplaceFormCode(lIndex As Integer, lOldCode As String, lNewCode As String)
Dim msg As String, msg2 As String, msg3 As String
msg = ReturnFormCode(lIndex)
If Len(msg) <> 0 Then
    msg2 = Replace(msg, lOldCode, lNewCode)
    TriggerSetFormEdit lIndex, msg2
    TriggerSetFormCode lIndex, msg2
End If
End Sub

Public Sub SetButtonCaption(lFormIndex As Integer, lButtonIndex As Integer, lCaption As String)
'MsgBox lFormIndex
'MsgBox lButtonIndex
lProject.SetButtonCaption lFormIndex, lButtonIndex, lCaption
End Sub

Public Sub SetButtonName(lFormIndex As Integer, lButtonIndex As Integer, lName As String)
lProject.SetButtonName lFormIndex, lButtonIndex, lName
End Sub

Public Sub SetModuleCursor(lModIndex As Integer, lPosition As Integer)
lProject.SetModuleCursor lModIndex, lPosition
End Sub

Public Sub SetButtonProporties(lFormIndex As Integer, lButtonIndex As Integer)
lProject.SetButtonProperties lFormIndex, lButtonIndex
End Sub

Public Function ReturnFormCursor(lIndex As Integer) As Integer
ReturnFormCursor = lProject.ReturnFormCursor(lIndex)
End Function

Public Sub SetFormCursor(lIndex As Integer, lPosition As Integer)
lProject.SetFormCursor lIndex, lPosition
End Sub

Public Sub TriggerSetFormEdit(lIndex As Integer, lCode As String)
lProject.SetFormCodeEdit lIndex, lCode
End Sub

Public Function ReturnFormCode(lIndex As Integer) As String
ReturnFormCode = lProject.ReturnFormCode(lIndex)
End Function

Public Sub TriggerSetFormCode(lIndex As Integer, lCode As String)
lProject.SetFormCode lIndex, lCode
End Sub

Public Sub AddObjectToForm(lIndex As Integer, lData As String)
If Len(lData) <> 0 And lIndex <> 0 Then
    lProject.AddObjectToForm lIndex, lData
End If
End Sub

Public Sub SetFormCode(lIndex As Integer, lCode As String)
lProject.SetFormCode lIndex, lCode
End Sub

Public Sub ShowFormCodeByIndex(lIndex As Integer)
lProject.ShowFormCodeByIndex lIndex
End Sub

Public Sub ShowFormCode(lFormName As String)
Dim i As Integer
For i = 0 To lProject.ReturnFormCount
'    MsgBox "!" & lProject.ReturnFormName(i) & "!"
    If Trim(lFormName) = Trim(lProject.ReturnFormName(i)) Then
    'If Trim(LCase(lFormName)) = Trim(LCase(lProject.ReturnFormName(i))) Then
        lProject.ShowFormCode i
        Exit Sub
    End If
Next i
End Sub

Public Sub NewProject(lName As String)
lProject.NewProject lName
End Sub

Public Function NewForm(lFormName As String) As Integer
NewForm = lProject.AddForm(lFormName)
End Function

Public Sub ToolBox(lShow As Boolean)
If lShow = True Then
    frmToolBox.Show
    mdiMain.picToolBox.Visible = True
Else
    Unload frmToolBox
    mdiMain.picToolBox.Visible = False
End If
End Sub

Public Sub Proporties(lShow As Boolean)
If lShow = True Then
    frmProporties.Show
    mdiMain.picProporties.Visible = True
Else
    Unload frmProporties
    mdiMain.picProporties.Visible = False
End If
End Sub

Public Sub ActivateResize()
MDIForm_Resize
End Sub

Public Function NewModule(lName As String, lFile As String, Optional lShow As Boolean, Optional lAddSubMain As Boolean) As Integer
NewModule = lProject.AddModule(lName, lFile, lShow, lAddSubMain)
End Function

Public Sub SetFocusForm(lIndex As Integer)
lProject.SetFormFocus lIndex
End Sub

Public Function ReturnProjectName() As String
ReturnProjectName = lProject.ReturnProjectName
End Function

Public Sub SetFormProperties(lFormIndex As Integer)
lProject.SetFormProperties lFormIndex
End Sub

Public Sub SetModuleCode(lIndex As Integer, lCode As String)
If lIndex <> 0 And Len(lCode) <> 0 Then lProject.SetModuleCode lIndex, lCode
End Sub

Public Sub SetModuleText(lIndex As Integer, lText As String, lVisible As Boolean)
If lIndex <> 0 And Len(lText) <> 0 Then
    lProject.SetModuleText lIndex, lText, lVisible
End If
End Sub

Public Sub ShowModule(lIndex As Integer)
lProject.ShowModule lIndex
End Sub

Public Sub SetModuleProperties(lCaption As String)
lProject.SetModuleProperties lCaption
End Sub

Private Sub ctlXPButton1_Click()

End Sub

Private Sub cmdCreateObject_Click()
If Len(txtCaption.Text) <> 0 And Len(txtName.Text) <> 0 Then
    Select Case cboObjectType.ListIndex
    Case 0
    Case 1
        ActiveForm.NewButton txtName.Text, txtCaption.Text
    Case 2
        ActiveForm.NewLabel txtName.Text, txtCaption.Text
    Case 8
        ActiveForm.NewSkinnedButton txtName.Text, txtCaption.Text
    End Select
End If
End Sub

Private Sub Image1_Click()

End Sub

Private Sub MDIForm_Load()
Dim b As Boolean, o As Boolean, msg As String
'ChangeWin mdiMain.hWnd, True, True, False, False, True, True, False, False, False, True, False
'frmProperties.Show
b = CBool(ReadINI(App.Path & "\DATA\CONFIG\UISETTINGS.INI", "SETTINGS", "AUTOMAXIMIZE", False))
mnuAutoMaximize.Checked = b
'Stop
If Len(lInterpreterLocation) = 0 Then
    lInterpreterLocation = App.Path & "\DATA\BIN\NVCDI.EXE"
    If DoesFileExist(lInterpreterLocation) = False Then
        MsgBox "Unable to locate interpreter! Exiting...", vbCritical
        End
    End If
End If
'lInterpreterLocation = "C:\Documents and Settings\Leon Aiossa\My Documents\My Projects\NVCDI\NVCDI.EXE"
Set lProject = New clsProject
frmToolBox.Show
frmProporties.Show
Me.Visible = True
'Proporties False
frmNewProject.Show 1
DoEvents
Me.Visible = True
End Sub

Public Function DoesSubExist(lIndex As Integer, lSub As String) As Boolean
DoesSubExist = lProject.DoesSubExist(lIndex, lSub)
End Function

Public Sub SetFocusModule(lIndex As Integer, lVisible As Boolean)
lProject.SetModuleFocus lIndex, lVisible
End Sub

Private Sub MDIForm_Resize()
On Local Error Resume Next
If Me.ScaleWidth > 2200 Then
    cboObject.Width = (Me.ScaleWidth - 800)
    'cboProporty.Width = Me.ScaleWidth / 10
    frmProporties.ctlPropertyList1.Height = mdiMain.ScaleHeight - 840
    frmProporties.ctlPropertyList1.ResizeControl 0, 0, frmProporties.Width, frmProporties.Height
'    Caption = "True"
'    cboObject.Width = (Me.ScaleWidth) - 4700
'    cboProporty.Left = cboObject.Left + cboObject.Width
'    If cboProporty.Width <> 4000 Then cboProporty.Width = 4000
Else
    Caption = "false"
End If
'frmProporties.Left = 7185
'frmProporties.Top = 345
'frmProporties.Height = 4695
'frmProporties.Width = 4000
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
WriteINI App.Path & "\DATA\CONFIG\UISETTINGS.INI", "SETTINGS", "AUTOMAXIMIZE", mnuAutoMaximize.Checked
ToolBox False
Proporties False
End
End Sub

Private Sub mnuAbout_Click()
'MsgBox App.Title & vbCrLf & "Version: " & App.Major & "." & App.Minor & vbCrLf & "Development environment" & vbCrLf & vbCrLf & "Team Nexgen thanks you for using " & App.Title & " as your development environment.", vbInformation
frmAbout.Show 1
End Sub

Private Sub mnuArrangeIcons_Click()
mdiMain.Arrange vbArrangeIcons
End Sub

Private Sub mnuAutoMaximize_Click()
If mnuAutoMaximize.Checked = False Then
    mnuAutoMaximize.Checked = True
Else
    mnuAutoMaximize.Checked = False
End If
End Sub

Private Sub mnuClose_Click()
frmProporties.tvwProject.Nodes.Clear
'Unload mdiMain.ActiveForm
End Sub

Private Sub mnuCompile_Click()
Dim mbox As VbMsgBoxResult
If lProject.ReturnProjectSaved = False Then
    mbox = MsgBox("Would you like to save this project now?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        lProject.SaveProject
    Else
        Exit Sub
    End If
End If
Shell lInterpreterLocation & " " & lProject.ReturnProjectFile
'MsgBox lInterpreterLocation & " " & lProject.ReturnProjectFile
End Sub

Private Sub mnuCut1_Click()
mdiMain.ActiveForm.SetCutObject
tlbTop.Buttons(7).Enabled = True
mdiMain.ActiveForm.HideHashMarks
End Sub

Private Sub mnuDelete_Click()
On Local Error Resume Next
Select Case ActiveForm.lCurrentObjectType
Case oFrame
    Unload ActiveForm.fraFrame(ActiveForm.lCurrentObjectIndex)
    ActiveForm.HideHashMarks
Case oTextBox
    Unload ActiveForm.txtText(ActiveForm.lCurrentObjectIndex)
    ActiveForm.HideHashMarks
Case oCommandButton
    Unload ActiveForm.cmdButton(ActiveForm.lCurrentObjectIndex)
    ActiveForm.HideHashMarks
Case oPictureBox
    Unload ActiveForm.picPicture(ActiveForm.lCurrentObjectIndex)
    ActiveForm.HideHashMarks
Case oLabel
    Unload ActiveForm.lblLabel(ActiveForm.lCurrentObjectIndex)
    ActiveForm.HideHashMarks
End Select
End Sub

Private Sub mnuDockToolBox_Click()
If mnuDockToolBox.Checked = False Then
    mnuDockToolBox.Checked = 1
    frmToolBox.DockToolBox
Else
    mnuDockToolBox.Checked = 0
    frmToolBox.UnDockToolBox
End If
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuForm_Click()
frmNewForm.Show 1
End Sub

Private Sub mnuGridLine_Click()
If mnuGridLine.Checked = True Then
    mnuGridLine.Checked = False
    WriteINI lProject.ReturnProjectFile, App.Title, "SnapToGridLines", "False"
Else
    mnuGridLine.Checked = True
    WriteINI lProject.ReturnProjectFile, App.Title, "SnapToGridLines", "False"
End If
End Sub

Private Sub mnuNewModule_Click()
frmNewModule.Show 1
End Sub

Private Sub mnuNewProject_Click()
frmNewProject.Show 1
End Sub

Private Sub mnuPaste1_Click()
mdiMain.ActiveForm.PasteCutObject
tlbTop.Buttons(7).Enabled = False
End Sub

Private Sub mnuProjectProporties_Click()
'frmCustomize.Show 1
End Sub

Private Sub mnuProperties_Click()
frmProporties.SetFocus
frmProporties.nTabControl1.SetFocus
frmProporties.nTabControl1.ActiveTab = 0
End Sub

Private Sub mnuSaveProject_Click()
lProject.SaveProject
End Sub

Private Sub mnuShowCodeEditor_Click()
frmEdit.Show
End Sub

Private Sub mnuSnap_Click()
If mnuSnap.Checked = True Then
    mnuSnap.Checked = False
    WriteINI lProject.ReturnProjectFile, App.Title, "Snap", "False"
Else
    mnuSnap.Checked = True
    WriteINI lProject.ReturnProjectFile, App.Title, "Snap", "True"
End If
End Sub

Private Sub mnuStart_Click()
Dim mbox As VbMsgBoxResult
If Len(lProject.ReturnProjectName) <> 0 Then
    If lProject.ReturnProjectSaved = False Then
        mbox = MsgBox("Would you like to save this project now?", vbYesNo + vbQuestion)
        If mbox = vbYes Then
            lProject.SaveProject
        Else
            Exit Sub
        End If
    End If
    Shell lInterpreterLocation & " " & lProject.ReturnProjectFile & " -run"
Else
    MsgBox "No project is active", vbExclamation
End If
End Sub

Private Sub mnuTileHorrizontal_Click()
mdiMain.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVerticle_Click()
mdiMain.Arrange vbTileVertical
End Sub

Private Sub mnuToggleProporties_Click()
If picProporties.Visible = True Then
    Proporties False
Else
    Proporties True
End If
End Sub

Private Sub mnuToggleToolBox_Click()
If mnuToggleToolBox.Checked = False Then
'If picToolBox.Visible = True Then
    ToolBox True
    mnuToggleToolBox.Checked = True
Else
    ToolBox False
    mnuToggleToolBox.Checked = False
End If
End Sub

Private Sub mnuViewCode_Click()
ActiveForm.ShowCode
End Sub

Private Sub tlbTop_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    frmNewProject.Show 1
Case 3
    lProject.SaveProject
Case 5
    mdiMain.ActiveForm.SetCutObject
    tlbTop.Buttons(7).Enabled = True
    mdiMain.ActiveForm.HideHashMarks
Case 7
    mdiMain.ActiveForm.PasteCutObject
    tlbTop.Buttons(7).Enabled = False
Case 12
    mnuStart_Click
End Select
End Sub
