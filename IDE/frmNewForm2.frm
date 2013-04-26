VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Fargo - About"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Fargo.ctlFrame ctlFrame1 
      Height          =   3390
      Left            =   0
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5980
      RoundedCorner   =   0   'False
      Caption         =   "About Fargo"
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
      Picture         =   "frmNewForm2.frx":0000
      ThemeColor      =   1
      ColorFrom       =   0
      ColorTo         =   0
      Begin Fargo.XPButton cmdClose 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   2880
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         MICON           =   "frmNewForm2.frx":0CDA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   120
         Picture         =   "frmNewForm2.frx":0CF6
         Stretch         =   -1  'True
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fargo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Developer: guideX"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Team Nexgen thanks you for using Project Fargo as your Development environment and hope you will use it in the future."
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0"
         Height          =   255
         Left            =   1320
         TabIndex        =   0
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   855
         Left            =   240
         Picture         =   "frmNewForm2.frx":24D8
         Stretch         =   -1  'True
         Top             =   600
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub ctlFrame1_Click()
FormDrag Me
End Sub

Private Sub ctlFrame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    FormDrag Me
End If
End Sub

Private Sub Form_Load()
Label2.Caption = "Version: " & App.Major & "." & App.Minor & App.Revision
End Sub

