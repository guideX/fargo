VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmEdit 
   Caption         =   "Visual Code Designer - Object (Code)"
   ClientHeight    =   3120
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   5355
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":03C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwCodeInsert 
      Height          =   2655
      Left            =   3480
      TabIndex        =   1
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   4683
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   0
      MousePointer    =   99
      MouseIcon       =   "frmEdit.frx":0714
   End
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4683
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmEdit.frx":0876
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Private lForm As Boolean
Private lSaved As Boolean
Private lFormIndex As Integer
Private lModIndex As Integer
Enum eCodeInsertTypes
    cForm = 1
    cModule = 2
End Enum

Public Sub ApplyCodeInsert(lValue As eCodeInsertTypes)
Dim msg As String
Select Case lValue
Case cForm
    msg = txtCode.Text
'    MsgBox msg
    tvwCodeInsert.Nodes.Clear
    tvwCodeInsert.Nodes.Add , , , "Primitives", 2
    tvwCodeInsert.Nodes(1).Bold = True
    'tvwCodeInsert.Nodes(1).ForeColor = vbBlue
    tvwCodeInsert.Nodes.Add 1, tvwChild, , "Form_Load", 1
    If InStr(msg, "Form_Load") <> 0 Then
        tvwCodeInsert.Nodes(2).ForeColor = vbBlack
    Else
        tvwCodeInsert.Nodes(2).ForeColor = vbCyan
    End If
    tvwCodeInsert.Nodes.Add 1, tvwChild, , "Form_Objects", 1
    If InStr(msg, "Form_Objects") <> 0 Then
        tvwCodeInsert.Nodes(3).ForeColor = vbBlack
    Else
        tvwCodeInsert.Nodes(3).ForeColor = vbRed
    End If
    tvwCodeInsert.Nodes.Add 1, tvwChild, , "Form_Click", 1
    If InStr(msg, "Form_Click") <> 0 Then
        tvwCodeInsert.Nodes(4).ForeColor = vbBlack
    Else
        tvwCodeInsert.Nodes(4).ForeColor = vbRed
    End If
    tvwCodeInsert.Nodes.Add 1, tvwChild, , "Form_DblClick", 1
    If InStr(msg, "Form_DblClick") <> 0 Then
        tvwCodeInsert.Nodes(5).ForeColor = vbBlack
    Else
        tvwCodeInsert.Nodes(5).ForeColor = vbRed
    End If
    tvwCodeInsert.Nodes(1).Expanded = True
Case cModule
End Select
End Sub

Public Sub SetModIndex(lIndex As Integer)
If lIndex <> 0 Then lModIndex = lIndex
End Sub

Public Sub SetFormIndex(lIndex As Integer)
If lIndex <> 0 Then
    lForm = True
    lFormIndex = lIndex
End If
End Sub

Public Function ReturnSaved() As Boolean
ReturnSaved = lSaved
End Function

Public Sub SetSaved(lValue As Boolean)
lSaved = lValue
End Sub

Private Sub cboObjects_Click()
'If LCase(cboObjects.Text) = "(general)" Then
'    cboProporties.Clear
'    cboProporties.AddItem "(Declarations)"
'    cboProporties.AddItem "Main"
'    cboProporties.ListIndex = 0
'End If
End Sub

Private Sub cboProporties_Click()
Dim msg() As String, i As Integer, n As Integer, msg2 As String, lLength As Long, msg3 As String
If Len(txtCode.Text) <> 0 Then
    msg = Split(txtCode.Text, vbCrLf)
    For i = 0 To UBound(msg)
        msg(i) = Trim(msg(i))
        lLength = lLength + Len(msg(i)) + 2
        'If LCase(Left(cboProporties.Text, 14)) = "(declarations)" Then
        '    txtCode.SelStart = 0
        '    txtCode.SetFocus
        '    Exit Sub
        'End If
        If LCase(Left(msg(i), 3)) = "sub" Then
            msg2 = Right(msg(i), Len(msg(i)) - 4)
            If InStr(msg2, "(") And InStr(msg2, ")") Then
                If InStr(msg2, "()") Then
                    msg2 = Replace(msg2, "()", "")
                Else
                    msg3 = Parse(msg2, "(", ")")
                End If
            End If
            'If LCase(msg2) = LCase(cboProporties.Text) Then
            '    txtCode.SelStart = lLength
            '    txtCode.SetFocus
            'End If
        End If
    Next i
End If
End Sub

Private Sub Command1_Click()
'frmFormCodeInsert.Show
'SetParent frmFormCodeInsert.hWnd, txtCode.hWnd
End Sub

Private Sub Form_GotFocus()
If mdiMain.mnuAutoMaximize.Checked = True Then Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
'CutRegion cboObjects.hWnd, cboObjects, True
'CutRegion cboProporties.hWnd, cboProporties, True
If mdiMain.mnuAutoMaximize.Checked = True Then Me.WindowState = vbMaximized
'txtCode.SelStart = mdiMain.ReturnFormCursor(lFormIndex)


End Sub

Private Sub Form_Resize()
On Local Error Resume Next
'cboObjects.Width = Me.ScaleWidth / 2
'cboProporties.Width = Me.ScaleWidth / 2
'cboProporties.Left = cboObjects.Width
Me.Icon = mdiMain.Icon
txtCode.Width = (Me.ScaleWidth - tvwCodeInsert.Width)
'txtCode.Height = Me.ScaleHeight - cboProporties.Height - 20
txtCode.Height = Me.ScaleHeight
If txtCode.Top <> 0 Then txtCode.Top = 0
If tvwCodeInsert.Top <> 0 Then tvwCodeInsert.Top = 0
tvwCodeInsert.Left = txtCode.Width
tvwCodeInsert.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel = 1
Me.Visible = False
If lForm = True Then
    mdiMain.SetFormCode lFormIndex, txtCode.Text
    mdiMain.SetFormCursor lFormIndex, txtCode.SelStart
Else
    mdiMain.SetModuleCode lModIndex, txtCode.Text
    mdiMain.SetModuleCursor lModIndex, txtCode.SelStart
    'MsgBox txtCode.SelLength
End If
End Sub

Public Sub SetTextFocus(lText As String)
On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer, n As Integer, t As Integer
StartMe:
msg = txtCode.Text
If Len(msg) <> 0 Then
    n = InStr(msg, lText)
    If n <> 0 Then
        msg = Right(msg, Len(msg) - n)
        t = n
        For i = 0 To Len(msg)
            msg2 = Left(msg, 1)
            msg = Right(msg, Len(msg) - 1)
            t = t + 1
            If Asc(msg2) = 13 Or Asc(msg2) = 10 Then
                txtCode.SelStart = (t + 1)
                txtCode.SetFocus
                Exit Sub
            End If
        Next i
    Else
        txtCode.Text = txtCode.Text & vbCrLf & "Sub " & tvwCodeInsert.SelectedItem.Text & "()" & vbCrLf & vbCrLf & "End Sub" & vbCrLf
        GoTo StartMe
    End If
End If

End Sub

Private Sub tvwCodeInsert_Click()
If tvwCodeInsert.SelectedItem.Text = "Primitives" Then Exit Sub
SetTextFocus tvwCodeInsert.SelectedItem.Text
ApplyCodeInsert cForm
End Sub

Private Sub txtCode_Change()
'mdiMain.SetModuleCode(
End Sub

Private Sub txtCode_GotFocus()
If mdiMain.mnuAutoMaximize.Checked = True Then Me.WindowState = vbMaximized
'mdiMain.SetModuleProperties Me.Caption
frmToolBox.EnableToolBox False
End Sub

Private Sub txtCode_KeyUp(KeyCode As Integer, Shift As Integer)
SetRTFColors txtCode, Me.hWnd
If lForm = True Then
    mdiMain.SetFormCode lFormIndex, txtCode.Text
Else
    mdiMain.SetModuleCode lModIndex, txtCode.Text
End If
End Sub
