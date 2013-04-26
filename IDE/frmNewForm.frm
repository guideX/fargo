VERSION 5.00
Begin VB.Form frmNewForm 
   BorderStyle     =   0  'None
   Caption         =   "New Form"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Fargo.ctlFrame ctlFrame1 
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2566
      RoundedCorner   =   0   'False
      Caption         =   "Create New Form"
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewForm.frx":000C
      ThemeColor      =   1
      ColorFrom       =   0
      ColorTo         =   0
      Begin Fargo.XPButton cmdCancel 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   960
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
         MICON           =   "frmNewForm.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtFormName 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   480
         Width           =   3615
      End
      Begin Fargo.XPButton cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   960
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
         MICON           =   "frmNewForm.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4680
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmNewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Len(txtFormName.Text) = 0 Then
    MsgBox "You did not specify a form name", vbExclamation
    txtFormName.SetFocus
    Beep
    Exit Sub
End If
mdiMain.NewForm txtFormName.Text
Unload Me
End Sub

Private Sub ctlFrame1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then FormDrag Me
End Sub

Private Sub Form_Load()
Me.Icon = mdiMain.Icon
End Sub
