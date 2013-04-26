VERSION 5.00
Begin VB.Form frmCompile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compile"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   1800
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtText 
      Height          =   495
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmCompile.frx":000C
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCaption 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "Enter caption here..."
      Top             =   4320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame 
      Caption         =   "Exe Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtExeFile 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "&Compile"
      Default         =   -1  'True
      Height          =   350
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblAddPic 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click to change picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   4800
      Width           =   1665
   End
   Begin VB.Image imgPic 
      Height          =   420
      Left            =   1560
      Picture         =   "frmCompile.frx":0021
      Stretch         =   -1  'True
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   4830
      Width           =   975
   End
End
Attribute VB_Name = "frmCompile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub chkPass_Click()
'On Local Error Resume Next
'If chkPass.Value > 0 Then
'    txtPass.Enabled = True
'    txtPass.SetFocus
'Else
'    txtPass.Enabled = False
'End If
'End Sub

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdCompile_Click()
If Len(txtExeFile.Text) <> 0 Then ReachCompiler Dir1.Path, txtExeFile.Text
End Sub

Private Sub Form_Load()
txtExeFile.Text = ReturnProj & ".exe"
Dir1.Path = App.Path
End Sub
