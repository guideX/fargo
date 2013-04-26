VERSION 5.00
Begin VB.Form frmForm 
   Caption         =   "Form"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBlank.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   Begin Fargo.ctlXPButton cmdSkinnedButton 
      DragMode        =   1  'Automatic
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Command0"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmBlank.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton optOption 
      Caption         =   "Option0"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrCheckResize 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   3240
      Top             =   1200
   End
   Begin VB.PictureBox shpMidBottom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   1920
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox shpMidRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   2160
      MousePointer    =   9  'Size W E
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox shpMidTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   1920
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox shpMidLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   1680
      MousePointer    =   9  'Size W E
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.CheckBox chkCheck 
      Appearance      =   0  'Flat
      Caption         =   "Check0"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmBlank.frx":0028
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox shpObjectBottomRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   2160
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox shpObjectTopRight 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   2160
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox shpObjectBottomLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   1680
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox shpObjectTopLeft 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   1680
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   100
   End
   Begin VB.PictureBox picPicture 
      Height          =   375
      Index           =   0
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Frame0"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Command0"
      Height          =   375
      Index           =   0
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape shpDrawborder 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      FillStyle       =   7  'Diagonal Cross
      Height          =   735
      Left            =   2760
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Label0"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private lObjectMax As Long
Private lLabelCount As Integer
Private lSkinnedButtonCount As Integer
Private lButtonCount As Integer
Private lFrameCount As Integer
Private lPictureCount As Integer
Private lTextBoxCount As Integer
Private lCheckBoxCount As Integer
Private lOptionButtonCount As Integer
Private lGrid As Integer
Private lObjectInMemory As Boolean
Private lCurrentObjectSet As Boolean
Private lFormIndex As Integer
Private lModIndex As Integer
'Private lDoubleClickEnabled As Boolean
Private lFormMouseDown As Boolean
'Private l

Enum eObjectType
    oForm = 0
    'oOther = 0
    oLabel = 1
    oPictureBox = 2
    oCommandButton = 3
    oOptionBox = 4
    oTextBox = 5
    oFrame = 6
    oCheckBox = 7
    oSkinnedButton = 8
End Enum
Private Enum eHashMarks
    hRight = 1
    hBottom = 2
    hLeft = 3
    hTop = 4
End Enum
Private lCurrentHashMark As eHashMarks
Public lCurrentObjectType As eObjectType
Private lCurrentObject As Object
Private lCutObject As Object
Public lCurrentObjectIndex As Integer
Private lDrawBorderStartingPointLeft As Integer
Private lDrawBorderStartingPointTop As Integer

Public Sub SetFormName(lIndex As Integer, lName As String)

End Sub

Public Function ReturnButtonCaption(lIndex As Integer)
ReturnButtonCaption = cmdButton(lIndex).Caption
End Function

Public Function ReturnButtonName(lIndex As Integer)
ReturnButtonName = cmdButton(lIndex).Tag
End Function

Public Function ReturnButtonWidth(lIndex As Integer)
ReturnButtonWidth = cmdButton(lIndex).Width
End Function

Public Function ReturnButtonHeight(lIndex As Integer)
ReturnButtonHeight = cmdButton(lIndex).Height
End Function

Public Function ReturnButtonLeft(lIndex As Integer)
ReturnButtonLeft = cmdButton(lIndex).Left
End Function

Public Function ReturnButtonTop(lIndex As Integer)
ReturnButtonTop = cmdButton(lIndex).Top
End Function

Public Function ReturnCurrentObjectIndex() As Integer
ReturnCurrentObjectIndex = lCurrentObjectIndex
End Function

Public Sub ShowCodeSub(lSub As String)

End Sub

Public Sub ShowCode()
mdiMain.ShowFormCodeByIndex lFormIndex
End Sub

Public Sub SetHashMarksOnCurrentObject()
If lObjectInMemory = True Then SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
End Sub

Public Sub SetModIndex(lIndex As Integer)
lModIndex = lIndex
End Sub

Public Sub SetFormIndex(lIndex As Integer)
lFormIndex = lIndex
End Sub

Public Sub PasteCutObject()
If lObjectInMemory = True Then
    lCutObject.Visible = True
    SetHashMarksOnCurrentObject
    lObjectInMemory = False
End If
End Sub

Public Sub SetCutObject()
If lCurrentObjectSet = True Then
    Set lCutObject = lCurrentObject
    lCutObject.Visible = False
    lObjectInMemory = True
End If
End Sub

Public Sub HideHashMarks()
shpObjectBottomLeft.Visible = False
shpObjectBottomRight.Visible = False
shpObjectTopLeft.Visible = False
shpObjectTopRight.Visible = False
shpMidBottom.Visible = False
shpMidLeft.Visible = False
shpMidRight.Visible = False
shpMidTop.Visible = False
End Sub

Public Sub NewOptionButton()
Dim o As OptionButton
If lOptionButtonCount < lObjectMax Then
    lOptionButtonCount = lOptionButtonCount + 1
    Set o = optOption(lOptionButtonCount)
    Load o
    o.Visible = True
    o.Caption = "Option" & Trim(Str(lOptionButtonCount))
    o.Value = False
    If lOptionButtonCount <> 1 Then
        o.Top = (o.Height * lOptionButtonCount) - o.Top
        o.Left = 100 * lOptionButtonCount
    End If
End If
End Sub

Public Sub NewCheckBox()
Dim c As CheckBox
If lCheckBoxCount < lObjectMax Then
    lCheckBoxCount = lCheckBoxCount + 1
    Set c = chkCheck(lCheckBoxCount)
    Load c
    c.Visible = True
    c.Caption = "Check" & Trim(Str(lCheckBoxCount))
    If lCheckBoxCount <> 1 Then
        c.Top = (c.Height * lCheckBoxCount) - c.Top
        c.Left = 100 * lCheckBoxCount
    End If
End If
End Sub

Public Sub NewPicture()
Dim p As PictureBox
'Stop
If lPictureCount < lObjectMax Then
    lPictureCount = lPictureCount + 1
    Set p = picPicture(lPictureCount)
    Load p
    p.Visible = True
    If lPictureCount <> 1 Then
        p.Top = (p.Height * lPictureCount) - p.Top
        p.Left = 100 * lPictureCount
    End If
End If
End Sub

Public Sub NewTextBox()
Dim t As TextBox
If lTextBoxCount < lObjectMax Then
    lTextBoxCount = lTextBoxCount + 1
    Set t = txtText(lTextBoxCount)
    Load t
    t.Visible = True
    t.Text = "Text" & Trim(Str(lPictureCount))
    If lTextBoxCount <> 1 Then
        t.Top = (t.Height * lTextBoxCount) - t.Top
        t.Left = 100 * lTextBoxCount
    End If
End If
End Sub

Public Sub NewFrame()
Dim F As Frame
If lFrameCount < lObjectMax Then
    lFrameCount = lFrameCount + 1
    Set F = fraFrame(lFrameCount)
    Load F
    F.Visible = True
    F.Caption = "Frame" & Trim(Str(lFrameCount))
    If lFrameCount <> 1 Then
        F.Top = (F.Height * lFrameCount) - F.Top
        F.Left = 100 * lFrameCount
    End If
End If
End Sub

Public Sub NewObject(lObject As eObjectType)
Select Case lObject
Case oPictureBox
    NewPicture
Case oCommandButton
    NewButton
Case oOptionBox
    NewOptionButton
Case oFrame
    NewFrame
Case oLabel
    NewLabel
End Select
End Sub

Sub ObjectDrag(Obj As Object, Optional lDontHash As Boolean, Optional lDoNotAlign As Boolean, Optional lResizeObjectRight As Boolean, Optional lResizeObjectLeft As Boolean, Optional lResizeObjectTop As Boolean, Optional lResizeObjectBottom As Boolean)
If lResizeObjectBottom = False And lResizeObjectLeft = False And lResizeObjectRight = False And lResizeObjectTop = False Then lCurrentHashMark = 0
tmrCheckResize.Enabled = True
ReleaseCapture
SendMessage Obj.hWnd, &HA1, 2, 0&
tmrCheckResize.Enabled = False
If mdiMain.mnuSnap.Checked = True And mdiMain.mnuGridLine.Checked = True And lDoNotAlign = False Then
    Obj.Left = AlignToGrid(Obj.Left)
    Obj.Top = AlignToGrid(Obj.Top)
End If
If lDontHash = False Then SetHashMarks Obj.Left, Obj.Top, Obj.Width, Obj.Height
If lResizeObjectRight = True Then
    lCurrentObject.Width = Obj.Left - lCurrentObject.Left
    SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
End If
If lResizeObjectLeft = True Then
    If tmrCheckResize.Enabled = False Then
        If Obj.Left < lCurrentObject.Left Then
            lCurrentObject.Width = lCurrentObject.Width + (lCurrentObject.Left - Obj.Left)
        Else
            lCurrentObject.Width = lCurrentObject.Width - (Obj.Left - lCurrentObject.Left)
        End If
        lCurrentObject.Left = Obj.Left
        SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
    End If
End If
If lResizeObjectTop = True Then
    If Obj.Top < lCurrentObject.Left Then
        lCurrentObject.Height = lCurrentObject.Height + (lCurrentObject.Top - Obj.Top)
    Else
        lCurrentObject.Height = lCurrentObject.Height - (Obj.Top - lCurrentObject.Top)
    End If
    lCurrentObject.Top = Obj.Top
    SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
End If
If lResizeObjectBottom = True Then
    lCurrentObject.Height = Obj.Top - lCurrentObject.Top
    SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
End If
Select Case lCurrentObjectType
Case oCommandButton
    'MsgBox "HEy    "
    'cmdButton_MouseDown lCurrentObjectIndex, 1, 0, 0, 0
End Select
End Sub

Function AlignToGrid(Value As Integer) As Integer
Dim i As Integer
For i = 0 To Value + lGrid Step lGrid
    If i > Value - (lGrid / 2) Then AlignToGrid = i: Exit Function
    If i > Value Then AlignToGrid = i - lGrid: Exit Function
Next i
End Function

Public Sub SetHashMarks(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
HideHashMarks
shpMidBottom.Top = lTop + lHeight
shpMidBottom.Left = (lLeft + lWidth / 2) - 60
shpMidTop.Left = (lLeft + lWidth / 2) - 60
shpMidTop.Top = lTop - 100
shpMidRight.Top = lTop + (lHeight / 2) - 60
shpMidRight.Left = lLeft + lWidth
shpMidLeft.Top = lTop + (lHeight / 2) - 60
shpMidLeft.Left = lLeft - 110
shpObjectTopLeft.Left = lLeft - 140
shpObjectTopLeft.Top = lTop - 140
shpObjectTopRight.Left = lLeft + lWidth
shpObjectTopRight.Top = lTop - 140
shpObjectBottomLeft.Left = lLeft - 140
shpObjectBottomLeft.Top = lTop + lHeight
shpObjectBottomRight.Top = lTop + lHeight
shpObjectBottomRight.Left = lLeft + lWidth
shpMidLeft.Visible = True
shpMidBottom.Visible = True
shpMidTop.Visible = True
shpMidRight.Visible = True
'shpObjectBottomRight.Visible = True
'shpObjectBottomLeft.Visible = True
'shpObjectTopLeft.Visible = True
'shpObjectTopRight.Visible = True
mdiMain.tlbTop.Buttons(5).Enabled = True
End Sub

'Public Sub NewSkinnedButton(lName As String, lCaption As String)
'Dim b As ctlXPButton, msg As String
'If Len(lName) <> 0 And Len(lCaption) <> 0 And lButtonCount < lObjectMax Then
'    lSkinnedButtonCount = lSkinnedButtonCount + 1
'    If Len(lName) = 0 Then lName = "SkinnedButton" & Trim(Str(lSkinnedButtonCount + 1))
'    If Len(lCaption) = 0 Then lCaption = "Skinned Button" & Trim(Str(lSkinnedButtonCount + 1))
'    Set b = cmdSkinnedButton(lSkinnedButtonCount)
'    Load b
'    If lButtonCount <> 1 Then
'        b.Top = (b. * lButtonCount) - b.Top
'        b.Left = 100 * lButtonCount
'    End If
'    b.Visible = True
'    b.Caption = lCaption
'    b.Tag = lName
'    msg = "Object SkinnedButton(" & lName & ", " & lCaption & ", " & b.Left & ", " & b.Top & ", " & b.Width & ", " & b.Height & ")"
'    If lFormIndex <> 0 Then mdiMain.AddObjectToForm lFormIndex, msg
'End If
'End Sub

Public Sub NewButton(Optional lName As String, Optional lCaption As String)
'Dim b As CommandButton, msg As String, msg2 As String, msg3 As String
Dim b As CommandButton, msg As String
If lButtonCount < lObjectMax Then
    lButtonCount = lButtonCount + 1
    If Len(lName) = 0 Then lName = "Button" & Trim(Str(lButtonCount + 1))
    If Len(lCaption) = 0 Then lCaption = "Button" & Trim(Str(lButtonCount + 1))
    Set b = cmdButton(lButtonCount)
    Load b
    If lButtonCount <> 1 Then
        b.Top = (b.Height * lButtonCount) - b.Top
        b.Left = 100 * lButtonCount
    End If
    b.Visible = True
    b.Caption = lCaption
    b.Tag = lName
    msg = "Object Button(" & lName & ", " & lCaption & ", " & b.Left & ", " & b.Top & ", " & b.Width & ", " & b.Height & ")"
    If lFormIndex <> 0 Then mdiMain.AddObjectToForm lFormIndex, msg
    mdiMain.SetFocus
    SetHashMarks cmdButton(lButtonCount).Left, cmdButton(lButtonCount).Top, cmdButton(lButtonCount).Width, cmdButton(lButtonCount).Height
    mdiMain.txtCaption.Text = cmdButton(lButtonCount).Caption
    mdiMain.txtName.Text = cmdButton(lButtonCount).Tag
    mdiMain.SetButtonProporties lFormIndex, lButtonCount
End If
End Sub

Public Sub NewLabel()
Dim l As Label
If lLabelCount < lObjectMax Then
    lLabelCount = lLabelCount + 1
    Set l = lblLabel(lLabelCount)
    Load l
    l.Visible = True
    l.Caption = "Label" & Trim(Str(lLabelCount))
    If lLabelCount <> 1 Then
        l.Top = (l.Height * lLabelCount) - l.Top
        l.Left = 100 * lLabelCount
    End If
End If
End Sub

Private Sub chkCheck_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set lCurrentObject = chkCheck(Index)
SetHashMarks chkCheck(Index).Left, chkCheck(Index).Top, chkCheck(Index).Width, chkCheck(Index).Height
If Button = 1 Then
    lCurrentObjectSet = True
    lCurrentObjectType = oCheckBox
    lCurrentObjectIndex = Index
    ObjectDrag chkCheck(Index), False, False, False, False, False, False
    mdiMain.txtName.Text = chkCheck(Index).Name
    mdiMain.txtCaption.Text = chkCheck(Index).Caption
Else
    PopupMenu mdiMain.mnuObjectMenu
End If
End Sub

Private Sub cmdButton_Click(Index As Integer)

'If lDoubleClickEnabled = True Then
'    mdiMain.ShowFormCodeByIndex lFormIndex
'    lDoubleClickEnabled = False
'    tmrCheckDoubleClick.Enabled = False
'End If
'If tmrCheckDoubleClick.Enabled = False Then
'    lDoubleClickEnabled = True
'    tmrCheckDoubleClick.Enabled = True
'End If
End Sub

Private Sub cmdButton_GotFocus(Index As Integer)
frmToolBox.EnableToolBox True
End Sub

Private Sub cmdButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As String, msg2 As String, msg3 As String, msg4 As String, i As Integer, F As Integer, msg5 As String ', msg6 As String
Set lCurrentObject = cmdButton(Index)
SetHashMarks cmdButton(Index).Left, cmdButton(Index).Top, cmdButton(Index).Width, cmdButton(Index).Height
'If Button = 1 Then
mdiMain.SetButtonProporties lFormIndex, Index
lCurrentObjectSet = True
lCurrentObjectType = oCommandButton
lCurrentObjectIndex = Index
msg = mdiMain.ReturnFormCode(lFormIndex)
msg2 = "Object Button(" & cmdButton(Index).Tag & ", " & cmdButton(Index).Caption & ", " & cmdButton(Index).Left & ", " & cmdButton(Index).Top & ", " & cmdButton(Index).Width & ", " & cmdButton(Index).Height & ")"
ObjectDrag cmdButton(Index), False, False, False, False, False, False
msg3 = "Object Button(" & cmdButton(Index).Tag & ", " & cmdButton(Index).Caption & ", " & cmdButton(Index).Left & ", " & cmdButton(Index).Top & ", " & cmdButton(Index).Width & ", " & cmdButton(Index).Height & ")"
msg4 = Replace(msg, msg2, msg3)
If msg4 = msg Then
    If InStr(msg4, cmdButton(Index).Tag) Then
        i = InStr(msg4, cmdButton(Index).Tag)
        If i <> 0 Then
            msg5 = Right(msg4, Len(msg4) - i + 1)
            msg3 = ""
            msg2 = msg5
            For F = 0 To Len(msg5)
                If Len(msg3) <> 0 Then
                    msg3 = msg3 & Left(msg2, 1)
                Else
                    msg3 = Left(msg2, 1)
                End If
                If Right(msg3, 1) = ")" Then Exit For
                If Len(msg2) <> 0 Then
                    msg2 = Right(msg2, Len(msg2) - 1)
                Else
                    Exit For
                End If
            Next F
            msg3 = "Object Button(" & msg3
            msg5 = "Object Button(" & cmdButton(Index).Tag & ", " & cmdButton(Index).Caption & ", " & cmdButton(Index).Left & ", " & cmdButton(Index).Top & ", " & cmdButton(Index).Width & ", " & cmdButton(Index).Height & ")"
            msg = Replace(msg4, msg3, msg5)
        End If
    Else
        mdiMain.AddObjectToForm lFormIndex, "Object Button(" & cmdButton(Index).Tag & ", " & cmdButton(Index).Caption & ", " & cmdButton(Index).Left & ", " & cmdButton(Index).Top & ", " & cmdButton(Index).Width & ", " & cmdButton(Index).Height & ")"
        Exit Sub
    End If
Else
    msg = msg4
End If
mdiMain.SetButtonProporties lFormIndex, Index
mdiMain.txtName.Text = cmdButton(Index).Tag
mdiMain.txtCaption.Text = cmdButton(Index).Caption
mdiMain.TriggerSetFormEdit lFormIndex, msg
mdiMain.TriggerSetFormCode lFormIndex, msg
frmProporties.SetObjectType oCommandButton
frmProporties.SetObjectIndex Index
frmProporties.SetWindIndex lFormIndex
Me.SetFocus
If Button = 2 Then
    PopupMenu mdiMain.mnuObjectMenu
End If
End Sub

'Private Sub cmdSkinnedButton_Click(Index As Integer)
'If lDoubleClickEnabled = True Then
'    mdiMain.ShowFormCodeByIndex lFormIndex
'    lDoubleClickEnabled = False
'    tmrCheckDoubleClick.Enabled = False
'End If
'If tmrCheckDoubleClick.Enabled = False Then
'    lDoubleClickEnabled = True
'    tmrCheckDoubleClick.Enabled = True
'End If
'End Sub

Private Sub cmdSkinnedButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set lCurrentObject = cmdSkinnedButton(Index)
SetHashMarks cmdSkinnedButton(Index).Left, cmdSkinnedButton(Index).Top, cmdSkinnedButton(Index).Width, cmdSkinnedButton(Index).Height
'cmdSkinnedButton(Index).Enabled = False
If Button = 1 Then
    Dim msg As String, msg2 As String, msg3 As String
    
    lCurrentObjectSet = True
    lCurrentObjectType = oSkinnedButton
    lCurrentObjectIndex = Index
    msg = mdiMain.ReturnFormCode(lFormIndex)
    msg2 = "Object SkinnedButton(" & cmdSkinnedButton(Index).Tag & ", " & cmdSkinnedButton(Index).Caption & ", " & cmdSkinnedButton(Index).Left & ", " & cmdSkinnedButton(Index).Top & ", " & cmdSkinnedButton(Index).Width & ", " & cmdSkinnedButton(Index).Height & ")"
    ObjectDrag cmdSkinnedButton(Index), False, False, False, False, False, False
    msg3 = "Object SkinnedButton(" & cmdSkinnedButton(Index).Tag & ", " & cmdSkinnedButton(Index).Caption & ", " & cmdSkinnedButton(Index).Left & ", " & cmdSkinnedButton(Index).Top & ", " & cmdSkinnedButton(Index).Width & ", " & cmdSkinnedButton(Index).Height & ")"
    msg = Replace(msg, msg2, msg3)
    mdiMain.txtName.Text = cmdSkinnedButton(Index).Caption
    mdiMain.txtCaption.Text = cmdSkinnedButton(Index).Name
    mdiMain.TriggerSetFormEdit lFormIndex, msg
    mdiMain.TriggerSetFormCode lFormIndex, msg
ElseIf Button = 2 Then
    PopupMenu mdiMain.mnuObjectMenu
End If
'cmdSkinnedButton(Index).Enabled = True
End Sub

Private Sub Command1_Click()
SetHashMarks 30, 30, 100, 100
End Sub

Private Sub Form_Click()
HideHashMarks
Set lCurrentObject = Nothing
'lCurrentObjectSet = False
'lCurrentObjectIndex = lFormIndex
frmProporties.SetObjectType oForm
frmProporties.SetObjectIndex lFormIndex
frmProporties.SetWindIndex lFormIndex
'MsgBox lFormIndex
mdiMain.tlbTop.Buttons(5).Enabled = False

End Sub

Private Sub Form_DblClick()
If lFormIndex <> 0 Then
    mdiMain.ShowFormCodeByIndex lFormIndex
End If
End Sub

Private Sub Form_GotFocus()
If mdiMain.mnuAutoMaximize.Checked = True Then Me.WindowState = vbMaximized
mdiMain.SetFormProperties lFormIndex
frmToolBox.EnableToolBox True
End Sub

Private Sub Form_Load()
If mdiMain.mnuAutoMaximize.Checked = True Then Me.WindowState = vbMaximized
Me.Icon = mdiMain.Icon
lblLabel(0).Left = 120
lblLabel(0).Top = 120
lObjectMax = 128
lGrid = 128
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As String, msg2 As String, msg3 As String
lFormMouseDown = True
If Button = 1 Then
    Select Case frmToolBox.GetCurrentToolBoxType
    Case oCommandButton
        HideHashMarks
        lDrawBorderStartingPointLeft = X / Screen.TwipsPerPixelX * 15
        lDrawBorderStartingPointTop = Y / Screen.TwipsPerPixelY * 15
        shpDrawborder.Left = lDrawBorderStartingPointLeft
        shpDrawborder.Top = lDrawBorderStartingPointTop
        shpDrawborder.Visible = True
    Case oForm
        mdiMain.SetFormProperties lFormIndex
        'mdiMain.SetButtonProporties lFormIndex, Index
        lCurrentObjectSet = False
        lCurrentObjectType = oForm
        lCurrentObjectIndex = 0
    End Select
ElseIf Button = 2 Then
    PopupMenu mdiMain.mnuFormMenu
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lFormMouseDown = True Then
    Select Case frmToolBox.GetCurrentToolBoxType
    Case oCommandButton
        Dim i As Integer, l As Integer
        shpDrawborder.Width = (X / Screen.TwipsPerPixelX * 15) - shpDrawborder.Left
        shpDrawborder.Height = (Y / Screen.TwipsPerPixelY * 15) - shpDrawborder.Top
    End Select
End If
Select Case frmToolBox.GetCurrentToolBoxType
Case oCommandButton
    If Me.MousePointer <> 2 Then Me.MousePointer = 2
Case Else
    If Me.MousePointer <> 0 Then Me.MousePointer = 0
End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As String, msg2 As String, msg3 As String
lFormMouseDown = False
shpDrawborder.Visible = False
Select Case frmToolBox.GetCurrentToolBoxType
Case oCommandButton
    Me.MousePointer = 0
    frmToolBox.SetCurrentToolBoxType 0
    frmToolBox.DeselectToolBoxTools
    NewButton
    msg2 = "Object Button(" & cmdButton(lButtonCount).Tag & ", " & cmdButton(lButtonCount).Caption & ", " & cmdButton(lButtonCount).Left & ", " & cmdButton(lButtonCount).Top & ", " & cmdButton(lButtonCount).Width & ", " & cmdButton(lButtonCount).Height & ")"
    cmdButton(lButtonCount).Width = shpDrawborder.Width
    cmdButton(lButtonCount).Height = shpDrawborder.Height
    cmdButton(lButtonCount).Left = shpDrawborder.Left
    cmdButton(lButtonCount).Top = shpDrawborder.Top
    msg3 = "Object Button(" & cmdButton(lButtonCount).Tag & ", " & cmdButton(lButtonCount).Caption & ", " & cmdButton(lButtonCount).Left & ", " & cmdButton(lButtonCount).Top & ", " & cmdButton(lButtonCount).Width & ", " & cmdButton(lButtonCount).Height & ")"
    mdiMain.SetButtonProporties lFormIndex, lButtonCount
    lCurrentObjectSet = True
    lCurrentObjectType = oCommandButton
    lCurrentObjectIndex = lButtonCount
    msg = mdiMain.ReturnFormCode(lFormIndex)
    msg = Replace(msg, msg2, msg3)
    mdiMain.txtName.Text = cmdButton(lButtonCount).Tag
    mdiMain.txtCaption.Text = cmdButton(lButtonCount).Caption
    mdiMain.TriggerSetFormEdit lFormIndex, msg
    mdiMain.TriggerSetFormCode lFormIndex, msg
    frmProporties.SetObjectType oCommandButton
    frmProporties.SetObjectIndex lButtonCount
    frmProporties.SetWindIndex lFormIndex
    cmdButton_MouseDown lButtonCount, 0, 0, 0, 0
    shpMidRight.Visible = True
End Select
End Sub

Private Sub Form_Paint()
Dim i As Integer, ii As Integer
For i = 0 To Me.ScaleWidth Step lGrid
    For ii = 0 To Me.ScaleHeight Step lGrid
        Me.PSet (i, ii), vbBlack
    Next
Next i
End Sub

Private Sub fraFrame_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set lCurrentObject = fraFrame(Index)
SetHashMarks fraFrame(Index).Left, fraFrame(Index).Top, fraFrame(Index).Width, fraFrame(Index).Height
If Button = 1 Then
    lCurrentObjectSet = True
    lCurrentObjectType = oFrame
    lCurrentObjectIndex = Index
    ObjectDrag fraFrame(Index), False, False, False, False, False, False
Else
    PopupMenu mdiMain.mnuObjectMenu
End If
End Sub

Private Sub lblLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set lCurrentObject = lblLabel(Index)
SetHashMarks lblLabel(Index).Left, lblLabel(Index).Top, lblLabel(Index).Width, lblLabel(Index).Height
If Button = 1 Then
    ObjectDrag lblLabel(Index), True, True, True, True, True, True
'    ObjectDrag lblLabel(Index), False, False, False, False, False, False
    lCurrentObjectSet = True
    lCurrentObjectType = oLabel
    lCurrentObjectIndex = Index
ElseIf Button = 2 Then
    PopupMenu mdiMain.mnuObjectMenu
End If
'With frmProporties.ColHead1
'    .Clear
'    .Add "Name", lblLabel(Index).Name
'    .Add "Alignment", lblLabel(Index).Alignment
'    .Add "Appearance", lblLabel(Index).Appearance
'    .Add "AutoSize", lblLabel(Index).AutoSize
'
'    .Add "BackColor", lblLabel(Index).BackColor
'    .Add "BackStyle", lblLabel(Index).BackStyle
'    .Add "BorderStyle", lblLabel(Index).BorderStyle
'    .Add "Caption", lblLabel(Index).Caption
'    .Add "Enabled", lblLabel(Index).Enabled
'    .Add "FontName", lblLabel(Index).Font.Name
'    .Add "FontSize", lblLabel(Index).Font.Size
'    .Add "ForeColor", lblLabel(Index).ForeColor
'    .Add "Height", lblLabel(Index).Height
'    .Add "Index", lblLabel(Index).Index
'    .Add "Left", lblLabel(Index).Left
'    .Add "MouseIcon", lblLabel(Index).MouseIcon
'    .Add "MousePointer", lblLabel(Index).MousePointer
'    .Add "RightToLeft", lblLabel(Index).RightToLeft
'    .Add "TabIndex", lblLabel(Index).TabIndex
'    .Add "Tag", lblLabel(Index).Tag
'    .Add "ToolTipText", lblLabel(Index).ToolTipText
'    .Add "Top", lblLabel(Index).Top
'    .Add "UseMnemonic", lblLabel(Index).UseMnemonic
'    .Add "Visible", lblLabel(Index).Visible
'    .Add "Width", lblLabel(Index).Width
'    .Add "WordWrap", lblLabel(Index).WordWrap
'    Call .Update
'End With
End Sub

Private Sub mnuDeleteObject_Click()

End Sub

Private Sub optOption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set lCurrentObject = optOption(Index)
SetHashMarks optOption(Index).Left, optOption(Index).Top, optOption(Index).Width, optOption(Index).Height
If Button = 1 Then
    ObjectDrag optOption(Index), False, False, False, False, False, False
    lCurrentObjectSet = True
    lCurrentObjectType = oOptionBox
    lCurrentObjectIndex = Index
ElseIf Button = 2 Then
    PopupMenu mdiMain.mnuObjectMenu
End If
End Sub

Private Sub picPicture_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Set lCurrentObject = picPicture(Index)
SetHashMarks picPicture(Index).Left, picPicture(Index).Top, picPicture(Index).Width, picPicture(Index).Height
If Button = 1 Then
    ObjectDrag picPicture(Index), False, False, False, False, False, False
    lCurrentObjectSet = True
    lCurrentObjectType = oPictureBox
    lCurrentObjectIndex = Index
    mdiMain.txtCaption.Text = ""
    mdiMain.txtName.Text = picPicture(Index).Name
ElseIf Button = 2 Then
    PopupMenu mdiMain.mnuObjectMenu
End If
End Sub

Private Sub shpMidBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lCurrentHashMark = hBottom
ObjectDrag shpMidBottom, True, True, False, False, False, True
End Sub

Private Sub shpMidLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lCurrentHashMark = hLeft
ObjectDrag shpMidLeft, True, True, False, True, False, False
End Sub

Private Sub shpMidLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lCurrentHashMark = 0
End Sub

Private Sub shpMidRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lCurrentHashMark = hRight
ObjectDrag shpMidRight, True, True, True, False, False, False
End Sub

Private Sub shpMidRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lCurrentHashMark = 0
End Sub

Private Sub shpMidTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lCurrentHashMark = hTop
ObjectDrag shpMidTop, True, True, False, False, True, False
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrCheckResize_Timer()
Select Case lCurrentHashMark
Case hRight
    lCurrentObject.Width = shpMidRight.Left - lCurrentObject.Left
    SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
Case hLeft
    If shpMidLeft.Left < lCurrentObject.Left Then
        lCurrentObject.Width = lCurrentObject.Width + (lCurrentObject.Left - shpMidLeft.Left)
    Else
        lCurrentObject.Width = lCurrentObject.Width - (shpMidLeft.Left - lCurrentObject.Left)
    End If
    lCurrentObject.Left = shpMidLeft.Left
    SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
Case hTop
    If shpMidTop.Top < lCurrentObject.Left Then
        lCurrentObject.Height = lCurrentObject.Height + (lCurrentObject.Top - shpMidTop.Top)
    Else
        lCurrentObject.Height = lCurrentObject.Height - (shpMidTop.Top - lCurrentObject.Top)
    End If
    lCurrentObject.Top = shpMidTop.Top
    SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
Case hBottom
    lCurrentObject.Height = shpMidBottom.Top - lCurrentObject.Top
    SetHashMarks lCurrentObject.Left, lCurrentObject.Top, lCurrentObject.Width, lCurrentObject.Height
End Select
End Sub

Private Sub txtText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
mdiMain.txtCaption.Text = txtText(Index).Text
mdiMain.txtName.Text = txtText(Index).Tag
Set lCurrentObject = txtText(Index)
SetHashMarks txtText(Index).Left, txtText(Index).Top, txtText(Index).Width, txtText(Index).Height
If Button = 1 Then
    ObjectDrag txtText(Index), False, False, False, False, False, False
    lCurrentObjectSet = True
    lCurrentObjectIndex = Index
    lCurrentObjectType = oTextBox
    mdiMain.txtName.Text = txtText(Index).Text
Else
    PopupMenu mdiMain.mnuObjectMenu
End If
End Sub
