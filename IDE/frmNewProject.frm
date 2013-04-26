VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmNewProject 
   BorderStyle     =   0  'None
   Caption         =   "Select Project"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewProject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Fargo.ctlFrame ctlFrame1 
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5318
      BackColor       =   0
      FillColor       =   0
      Style           =   4
      RoundedCorner   =   0   'False
      Caption         =   "Create a new project"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewProject.frx":000C
      ThemeColor      =   5
      ColorFrom       =   12632256
      ColorTo         =   8421504
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   2175
         Width           =   4215
      End
      Begin MSComctlLib.ListView lvwNewProject 
         Height          =   1545
         Left            =   105
         TabIndex        =   3
         Top             =   525
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2725
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6174
         EndProperty
      End
      Begin Fargo.XPButton cmdCancel 
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Cancel"
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
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmNewProject.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Fargo.XPButton cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "OK"
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
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmNewProject.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line l 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   0
         X2              =   6000
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Line l 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   6000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line l 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   -360
         X2              =   5640
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Line l 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   -360
         X2              =   5640
         Y1              =   2070
         Y2              =   2070
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Project Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2190
         Width           =   1815
      End
   End
   Begin VB.Frame fraNewProject 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1540
      Left            =   6960
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
   End
End
Attribute VB_Name = "frmNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lIndex As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdNewProject_Click()
fraNewProject.Visible = True
'fraOpenProject.Visible = False
txtName.SetFocus
End Sub

Private Sub cmdOK_Click()
Dim msg As String, i As Integer, n As Integer
msg = txtName.Text
If Len(msg) <> 0 Then
    msg = Replace(msg, " ", "_")
    If InStr(msg, "$") Or InStr(msg, ":") Or InStr(msg, "/") Or InStr(msg, "\") Or InStr(msg, ",") Or InStr(msg, " ") Or InStr(msg, ".") Then
        MsgBox "Invalid Project Name", vbExclamation
        Exit Sub
    End If
    Select Case lvwNewProject.SelectedItem.Index
    Case 1
        Unload Me
        mdiMain.NewProject msg
        i = mdiMain.NewForm("frm" & msg)
        n = mdiMain.NewModule("mdl" & msg, "mdl" & msg & ".txt", False, True)
        mdiMain.SetModuleText n, "Sub Main()" & vbCrLf & "frm" & msg & ".Show" & vbCrLf & "End Sub" & vbCrLf, False
    Case 1
        Unload Me
        mdiMain.NewProject msg
        frmNewModule.chkSubMain.Value = 1
        frmNewModule.chkSubMain.Enabled = False
        frmNewModule.Show 1
    Case 2
        mdiMain.NewProject msg
        frmNewForm.Show 1
    End Select
    Unload Me
Else
    MsgBox "Specify a project name", vbExclamation
End If
End Sub

Private Sub cmdOpenProject_Click()
fraNewProject.Visible = False
'fraOpenProject.Visible = True
txtName.SetFocus
End Sub

Private Sub ctlFrame1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
FormDrag Me
End Sub

Private Sub Form_Load()
On Local Error Resume Next
lvwNewProject.ListItems.Add , , "Win32_Standard"
lvwNewProject.ListItems(1).SubItems(1) = "Project with a window and module"
lvwNewProject.ListItems.Add , , "Win32_Blank"
lvwNewProject.ListItems(2).SubItems(1) = "Project with a blank module"
lvwNewProject.ListItems.Add , , "Win32_Windowed"
lvwNewProject.ListItems(3).SubItems(1) = "Project with a window"
Me.Icon = mdiMain.Icon
txtName.SetFocus
End Sub

Private Sub lvwNewProject_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
txtName.SetFocus
End Sub
