VERSION 5.00
Begin VB.Form frmNewModule 
   BorderStyle     =   0  'None
   Caption         =   "New Module"
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewModule.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Fargo.ctlFrame ctlFrame1 
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3201
      RoundedCorner   =   0   'False
      Caption         =   "New Module"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmNewModule.frx":000C
      ThemeColor      =   1
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.CheckBox chkSubMain 
         Caption         =   "Add 'Sub Main()'"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   915
         Width           =   1575
      End
      Begin VB.TextBox txtModuleName 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   600
         Width           =   3615
      End
      Begin Fargo.ctlXPButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "Cancel"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmNewModule.frx":0CE6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Fargo.ctlXPButton cmdOK 
         Default         =   -1  'True
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   14
         TX              =   "OK"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmNewModule.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4680
         Y1              =   1215
         Y2              =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   4680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmNewModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim msg As String, b As Boolean
msg = txtModuleName.Text
If chkSubMain.Value = 1 Then
    b = True
Else
    b = False
End If
Unload Me
mdiMain.NewModule msg, msg & ".txt", True, b
End Sub

Private Sub ctlFrame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_Load()
Me.Icon = mdiMain.Icon
End Sub
