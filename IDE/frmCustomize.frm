VERSION 5.00
Object = "{004BB5D5-7D55-4298-B546-5ABB5C35F3AC}#1.0#0"; "NexgenTab.ocx"
Begin VB.Form frmCustomize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customize"
   ClientHeight    =   2760
   ClientLeft      =   3135
   ClientTop       =   3270
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6375
   Begin nTab.nTabControl nTabControl1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
      TabCaption(0)   =   "General"
      TabContCtrlCnt(0)=   4
      Tab(0)ContCtrlCap(1)=   "cboStartupObject"
      Tab(0)ContCtrlCap(2)=   "txtName"
      Tab(0)ContCtrlCap(3)=   "Label2"
      Tab(0)ContCtrlCap(4)=   "Label1"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabStyle        =   1
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      Begin VB.ComboBox cboStartupObject 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Startup Object:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCustomize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
