VERSION 5.00
Begin VB.UserControl ctlPropertyList 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3990
   ScaleWidth      =   4110
   Begin VB.PictureBox picStuff 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   3735
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.TextBox txtValue 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   1920
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.PictureBox picDropDown 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3600
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.ListBox lstDropDown 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   1620
         IntegralHeight  =   0   'False
         ItemData        =   "ctlPropertyList.ctx":0000
         Left            =   2160
         List            =   "ctlPropertyList.ctx":0002
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line lnVerticle 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line lnHorrizontal 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   3840
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.PictureBox picDropDown2 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   120
      Picture         =   "ctlPropertyList.ctx":0004
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox picDropDown1 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   360
      Picture         =   "ctlPropertyList.ctx":02AE
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.VScrollBar ctlScroll 
      Height          =   3975
      LargeChange     =   10
      Left            =   3840
      Max             =   100
      SmallChange     =   3
      TabIndex        =   0
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "ctlPropertyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lVar As Integer
Private lTextBoxHeight As Long
Private lTextBoxColor As ColorConstants
Private lPropTypes(128) As ePropTypes
Enum ePropTypes
    nText = 1
    nDropDown = 2
    'nColorPicker = 3
    'nFontSelector = 4
End Enum
Public Event ValueChanged(lIndex As Integer, lName As String, lValue As String)
Public Event KeyPress(lKey As Integer)

Private Function DoesPropertyIndexExist(lIndex As Integer) As Boolean
On Local Error GoTo ErrHandler
If lIndex <> 0 Then
    If Len(txtName(lIndex).Text) <> 0 Then
        DoesPropertyIndexExist = True
    Else
        DoesPropertyIndexExist = True
    End If
End If
Exit Function
ErrHandler:
    Exit Function
End Function

Private Sub AddPropertyIndex(lIndex As Integer)
On Local Error Resume Next
Load txtName(lIndex)
Load txtValue(lIndex)
Load lnHorrizontal(lIndex)
If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub ClearPropList()
Dim i As Integer
For i = 1 To txtName.Count - 1
    txtName(i).Text = ""
    txtName(i).Visible = False
    txtValue(i).Text = ""
    txtValue(i).Visible = False
    lnHorrizontal(i).Visible = False
Next i
End Sub

Public Sub SetTextBoxOverColor(lcolor As ColorConstants)
lTextBoxColor = lcolor
End Sub

Public Sub SetTextBoxHeight(lHeight As Long)
lTextBoxHeight = lHeight
End Sub

Public Function ReturnPropertyIndex(lName As String) As Integer
Dim i As Integer
For i = 0 To txtName.Count
    If LCase(Trim(lName)) = LCase(Trim(txtName(i).Text)) Then
        ReturnPropertyIndex = i
       Exit For
    End If
Next i
End Function

Public Function ReturnPropertyValue(lIndex As Integer) As String
ReturnPropertyValue = txtValue(lIndex).Text
End Function

Public Function ReturnPropertyValueByName(lName As String) As String
ReturnPropertyValueByName = txtValue(ReturnPropertyIndex(lName)).Text
End Function

Public Sub SetPropertyValue(lIndex As Integer, lValue As String)
txtValue(lIndex).Text = lValue
End Sub

Public Sub SetPropertyValueByName(lName As String, lValue As String)
txtValue(ReturnPropertyIndex(lName)).Text = lValue
End Sub

Public Sub ResizeControl(lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer)
lVar = 0
ctlScroll.Left = lWidth - ctlScroll.Width
ctlScroll.Height = lHeight
lnVerticle.X1 = txtValue(0).Width + 20
lnVerticle.X2 = txtValue(0).Width + 20
lnVerticle.Y2 = lHeight
picDropDown.Visible = False
lstDropDown.Visible = False
picStuff.Width = lWidth - ctlScroll.Width
picStuff.Height = lHeight
For lVar = 0 To txtName.Count - 1
    txtName(lVar).Width = lWidth / 2
    txtValue(lVar).Width = (lWidth / 2) - 20
    txtValue(lVar).Left = txtName(lVar).Width + 20
    lnHorrizontal(lVar).X2 = lWidth
Next lVar
End Sub

Public Sub AddProporty(lName As String, lProperty As String, lType As ePropTypes, Optional lDropDownItems As String)
Dim i As Integer, m As Integer, p As Integer, b As Boolean
If lTextBoxHeight = 0 Then lTextBoxHeight = 200
i = txtName.Count
p = i - 1
If p <> 0 Then
    If DoesPropertyIndexExist(p) = False Then
        MsgBox "Previous Property doesn't exist"
        Stop
    Else
        Do Until b = True
        
            If p - 1 <> 0 Then
                p = p - 1
                If Len(txtName(p).Text) = 0 Then
                    i = p
                Else
                    b = True
                End If
            Else
                b = True
            End If
        Loop
    End If
End If
If DoesPropertyIndexExist(i) = False Then AddPropertyIndex i
txtName(i).Height = lTextBoxHeight
txtName(i).Visible = True
txtName(i).Text = lName
txtValue(i).Height = lTextBoxHeight
txtValue(i).Visible = True
txtValue(i).Text = lProperty
m = txtName(i).Height + 15
txtName(i).Top = ((txtName(i).Height + 10) * i) - m
txtValue(i).Top = ((txtValue(i).Height + 10) * i) - m
lnHorrizontal(i).Visible = True
lnHorrizontal(i).Y1 = txtName(i).Top + txtName(i).Height - m
lnHorrizontal(i).Y2 = txtName(i).Top + txtName(i).Height - m
lPropTypes(i) = lType
Select Case lType
Case nDropDown
    If Len(lDropDownItems) <> 0 Then txtValue(i).Tag = lDropDownItems
End Select
End Sub

Private Sub ctlScroll_Change()
picStuff.Top = picStuff.Height / 100 - ctlScroll.Value * 20
picStuff.Height = picStuff.Height + 1000
End Sub

Private Sub ctlScroll_Scroll()
picStuff.Top = picStuff.Height / 100 - ctlScroll.Value * 20
picStuff.Height = ctlScroll.Height - picStuff.Top
End Sub

Private Sub lstDropDown_Click()
If Len(lstDropDown.Text) <> 0 Then
    txtValue(Int(lstDropDown.Tag)).Text = lstDropDown.Text
    lstDropDown.Visible = False
End If
End Sub

Private Sub picDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picDropDown.Picture = picDropDown2.Picture
End Sub

Private Sub picDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
picDropDown.Picture = picDropDown1.Picture
If lstDropDown.Visible = True Then
    lstDropDown.Visible = False
Else
    lstDropDown.Visible = True
End If
End Sub

Private Sub txtName_GotFocus(Index As Integer)
If lTextBoxColor = 0 Then
    txtName(Index).BackColor = &H8000000D
    txtName(Index).ForeColor = vbWhite
Else
    txtName(Index).BackColor = lTextBoxColor
    txtName(Index).ForeColor = vbBlack
End If
picDropDown.Visible = False
lstDropDown.Visible = False
End Sub

Private Sub txtName_LostFocus(Index As Integer)
txtName(Index).BackColor = vbWhite
txtName(Index).ForeColor = vbBlack
End Sub

Private Sub txtValue_Change(Index As Integer)
RaiseEvent ValueChanged(Index, txtName(Index).Text, txtValue(Index).Text)
End Sub

Private Sub txtValue_Click(Index As Integer)
If lPropTypes(Index) = nDropDown Then picDropDown.Visible = True
picDropDown.Top = txtValue(Index).Top
picDropDown.Left = (txtValue(Index).Width * 2) - picDropDown.Width * 2
End Sub

Private Sub txtValue_GotFocus(Index As Integer)
Dim msg() As String, i As Integer
If lTextBoxColor = 0 Then
    txtName(Index).BackColor = &H8000000D
    txtName(Index).ForeColor = vbWhite
Else
    txtName(Index).BackColor = lTextBoxColor
    txtName(Index).ForeColor = vbBlack
End If
txtValue(Index).SelStart = 0
txtValue(Index).SelLength = Len(txtValue(Index).Text)
If lPropTypes(Index) = nDropDown Then
    msg = Split(txtValue(Index).Tag, vbCrLf)
    lstDropDown.Clear
    For i = 0 To UBound(msg)
        If Len(msg(i)) <> 0 Then lstDropDown.AddItem msg(i)
    Next i
    picDropDown.Top = txtValue(Index).Top
    picDropDown.Left = (txtValue(Index).Width * 2) - picDropDown.Width * 2
    lstDropDown.Left = txtValue(Index).Left
    lstDropDown.Top = txtValue(Index).Top + txtValue(Index).Height
    lstDropDown.Width = txtName(Index).Width - 270
    lstDropDown.Tag = Trim(Str(Index))
    picDropDown.Visible = True
    If lPropTypes(Index) <> nDropDown Then
        picDropDown.Visible = False
        lstDropDown.Visible = False
    End If
Else
    lstDropDown.Visible = False
    picDropDown.Visible = False
End If
End Sub

Private Sub txtValue_KeyPress(Index As Integer, KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
If KeyAscii = 13 Then
    KeyAscii = 0
    mdiMain.ActiveForm.SetFocus
End If
End Sub

Private Sub txtValue_LostFocus(Index As Integer)
txtName(Index).BackColor = vbWhite
txtName(Index).ForeColor = vbBlack
End Sub

Private Sub UserControl_Initialize()
picDropDown.Picture = picDropDown1.Picture
End Sub
