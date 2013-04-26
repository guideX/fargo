VERSION 5.00
Begin VB.Form frmNothing 
   Caption         =   "Nothing"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNothing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   2400
   Visible         =   0   'False
   Begin VB.CommandButton cmdButton 
      Caption         =   "Nothing"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrExecuteForm_Load 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmNothing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lButtonCount As Integer
Private lFormIndex As Integer

Public Sub SetFormIndex(lIndex As Integer)
On Local Error Resume Next
lFormIndex = lIndex
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormIndex(lIndex As Integer)"
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Local Error Resume Next
Trigger_cmdButton_Click Trim(lFormIndex), Index, cmdButton(Index).Tag
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub cmdButton_Click(Index As Integer)"
End Sub

Private Sub cmdButton_GotFocus(Index As Integer)
On Local Error Resume Next
Trigger_cmdButton_GotFocus Trim(lFormIndex), Index, cmdButton(Index).Tag
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub cmdButton_GotFocus(Index As Integer)"
End Sub

Private Sub cmdButton_LostFocus(Index As Integer)
On Local Error Resume Next
Trigger_cmdButton_LostFocus Trim(lFormIndex), Index, cmdButton(Index).Tag
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub cmdButton_LostFocus(Index As Integer)"
End Sub

Private Sub Form_Activate()
On Local Error Resume Next
If lFormIndex <> 0 Then
    Trigger_Form_Activate lFormIndex
    Trigger_Form_Position lFormIndex
End If
End Sub

Private Sub Form_Click()
On Local Error Resume Next
If lFormIndex <> 0 Then Trigger_Form_Click Trim(lFormIndex)
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub Form_Click()"
End Sub

Private Sub Form_DblClick()
On Local Error Resume Next
If lFormIndex <> 0 Then Trigger_Form_DblClick Trim(lFormIndex)
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub Form_DblClick()"
End Sub

Private Sub Form_GotFocus()
On Local Error Resume Next
If Len(lFormIndex) <> 0 Then Trigger_Form_GotFocus Trim(lFormIndex)
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub Form_GotFocus()"
End Sub

Private Sub Form_Load()
On Local Error Resume Next
tmrExecuteForm_Load.Enabled = True
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub Form_Load()"
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
If Len(lFormIndex) <> 0 Then Trigger_Form_Resize Trim(lFormIndex)
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub Form_Resize()"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
If Len(lFormIndex) <> 0 Then Trigger_Form_Unload Trim(lFormIndex)
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub Form_Unload(Cancel As Integer)"
End Sub

Private Sub tmrExecuteForm_Load_Timer()
On Local Error Resume Next
If lFormIndex <> 0 Then
    Trigger_Form_Load Trim(lFormIndex)
    Trigger_Form_Objects Trim(lFormIndex)
    Me.Visible = True
End If
tmrExecuteForm_Load.Enabled = False
If Err.Number <> 0 Then ProcessRuntimeError Err.Number, Err.Description, "Private Sub tmrExecuteForm_Load_Timer()"
End Sub
