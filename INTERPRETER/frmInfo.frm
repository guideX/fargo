VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compiled Information"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtModule 
      Height          =   1215
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.ListBox lstModules 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdRun_Click()
'ManSetSubMain PropBag.ReadProperty("SUBMAIN")
'RunProject
End Sub

Private Sub lstModules_Click()
'txtModule.Text = PropBag.ReadProperty(lstModules.Text)
End Sub
