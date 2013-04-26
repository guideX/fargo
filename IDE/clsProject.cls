VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Enum eProjectTypes
'    pEmptyProject = 0
    pExe = 1
'    pOCX = 2
'    pDLL = 3
End Enum
Private Type gModules
    mName As String
    mFileName As String
    mWindow As New frmEdit
    mCode As String
    mEditCursor As Integer
End Type
Private Type gForm
    fName As String
    fForm As frmForm
    fEdit As frmEdit
    fEditCursor As Integer
    fFileName As String
    fCode As String
End Type
Private Type gProject
    pProjectSaved As Boolean
    pProjectFile As String
    pName As String
    pForms(128) As gForm
    pFormCount As Integer
    pModuleCount As Integer
    pModules(128) As gModules
    'pClass(128) As gClasses
    pType As eProjectTypes
    pIcon As String
End Type
Private lProject As gProject

Public Sub SetButtonCaption(lFormIndex As Integer, lButtonIndex As Integer, lCaption As String)
Dim msg As String, msg2 As String, msg3 As String
With lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex)
    msg = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    msg2 = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & lCaption & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    mdiMain.ReplaceFormCode lFormIndex, msg, msg2
    .Caption = lCaption
End With
End Sub

Public Sub SetButtonName(lFormIndex As Integer, lButtonIndex As Integer, lName As String)
Dim msg As String, msg2 As String
With lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex)
    msg = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    msg2 = "Object Button(" & Trim(lName) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    mdiMain.ReplaceFormCode lFormIndex, msg, msg2
    .Tag = lName
End With
'lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex).Tag = lName
End Sub

Public Sub SetButtonLeft(lFormIndex As Integer, lButtonIndex As Integer, lLeft As Integer)
Dim msg As String, msg2 As String
With lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex)
    msg = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    msg2 = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & lLeft & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    mdiMain.ReplaceFormCode lFormIndex, msg, msg2
    .Left = lLeft
    mdiMain.ActiveForm.SetHashMarks .Left, .Top, .Width, .Height
End With
End Sub

Public Sub SetButtonTop(lFormIndex As Integer, lButtonIndex As Integer, lTop As Integer)
Dim msg As String, msg2 As String
With lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex)
    msg = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    msg2 = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & lTop & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    mdiMain.ReplaceFormCode lFormIndex, msg, msg2
    .Top = lTop
    mdiMain.ActiveForm.SetHashMarks .Left, .Top, .Width, .Height
End With
End Sub

Public Sub SetButtonWidth(lFormIndex As Integer, lButtonIndex As Integer, lWidth As Integer)
Dim msg As String, msg2 As String
With lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex)
    msg = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    msg2 = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & lWidth & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    mdiMain.ReplaceFormCode lFormIndex, msg, msg2
    .Width = lWidth
    mdiMain.ActiveForm.SetHashMarks .Left, .Top, .Width, .Height
End With
End Sub

Public Sub SetButtonHeight(lFormIndex As Integer, lButtonIndex As Integer, lHeight As Integer)
Dim msg As String, msg2 As String
With lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex)
    msg = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonHeight(lButtonIndex) & ")"
    msg2 = "Object Button(" & Trim(mdiMain.ActiveForm.ReturnButtonName(lButtonIndex)) & ", " & Trim(mdiMain.ActiveForm.ReturnButtonCaption(lButtonIndex)) & ", " & mdiMain.ActiveForm.ReturnButtonLeft(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonTop(lButtonIndex) & ", " & mdiMain.ActiveForm.ReturnButtonWidth(lButtonIndex) & ", " & lHeight & ")"
    mdiMain.ReplaceFormCode lFormIndex, msg, msg2
    .Height = lHeight
    mdiMain.ActiveForm.SetHashMarks .Left, .Top, .Width, .Height
End With
End Sub

Public Sub SetFormCursor(lIndex As Integer, lCursor As Integer)
lProject.pForms(lIndex).fEditCursor = lCursor
End Sub

Public Function ReturnFormCursor(lIndex As Integer)
ReturnFormCursor = lProject.pForms(lIndex).fEditCursor
End Function

Public Function ReturnFormCode(lIndex As Integer) As String
ReturnFormCode = lProject.pForms(lIndex).fCode
End Function

Public Sub SetModuleText(lIndex As Integer, lText As String, lVisible As Boolean)
If Len(lText) <> 0 Then
    lProject.pModules(lIndex).mWindow.txtCode.Text = lText
    lProject.pModules(lIndex).mWindow.Visible = lVisible
End If
End Sub

Public Sub OpenProject(lProjectFile As String)
Dim i As Integer, lNode As Node
With lProject
    .pName = ReadINI(lProjectFile, "PROJECT", "NAME", "")
    .pIcon = ReadINI(lProjectFile, "PROJECT", "ICON", "")
    .pType = ReadINI(lProjectFile, "PROJECT", "TYPE", 0)
    .pProjectFile = lProjectFile
    .pFormCount = ReadINI(lProjectFile, "PROJECT", "MODULES", 0)
    .pModuleCount = ReadINI(lProjectFile, "PROJECT", "FORMS", 0)
    For i = 1 To .pFormCount
        .pForms(i).fFileName = ReadINI(lProjectFile, "FORM" & Trim(Str(i)), "FILE", "")
        .pForms(i).fName = ReadINI(lProjectFile, "FORM" & Trim(Str(i)), "NAME", "")
        Set .pForms(i).fForm = New frmForm
        Set .pForms(i).fEdit = New frmEdit
        .pForms(i).fForm.Visible = False
        .pForms(i).fForm.Caption = ReadINI(lProjectFile, "FORM" & Trim(Str(i)), "CAPTION", "")
        .pForms(i).fEdit.Caption = .pForms(i).fFileName
        .pForms(i).fEdit.txtCode.Text = ReadFile(Left(lProject.pProjectFile, Len(lProject.pProjectFile) - Len(GetFileTitle(lProject.pProjectFile))) & .pForms(i).fFileName)
        .pForms(i).fEdit.Visible = False
        Set lNode = frmProporties.tvwProject.Nodes.Add(3, tvwChild, , .pForms(i).fName, 1)
        lNode.Tag = lProject.pFormCount
        frmProporties.tvwProject.Nodes(3).Expanded = True
    Next i
    For i = 1 To .pModuleCount
        .pModules(i).mFileName = ReadINI(lProjectFile, "MODULE" & Trim(Str(i)), "FILE", "")
        .pModules(i).mName = ReadINI(lProjectFile, "MODULE" & Trim(Str(i)), "NAME", "")
        Set .pModules(i).mWindow = New frmEdit
        '.pModules(I).mWindow.txtCode.Text = ReadFile(
    Next i
End With
End Sub

Public Sub NewProject(lProjectName As String)
Dim lNode As Node
If Len(lProjectName) <> 0 Then
    frmProporties.tvwProject.Nodes.Clear
    Set lNode = frmProporties.tvwProject.Nodes.Add(, , lProjectName, lProjectName, 3)
    frmProporties.tvwProject.Nodes.Add 1, tvwChild, , "Modules", 2
    frmProporties.tvwProject.Nodes.Add 1, tvwChild, , "Forms", 2
    frmProporties.tvwProject.Nodes(1).Expanded = True
    ClearProject
    lProject.pName = lProjectName
    frmProporties.FormDragger1.Caption = "Project - " & lProjectName
    mdiMain.Caption = lProjectName & " - " & App.Title & " (Designer)"
    SetProjectType pExe
End If
End Sub

Public Sub SetProjectType(lType As eProjectTypes)
lProject.pType = lType
End Sub

Public Function ReturnProjectFile() As String
ReturnProjectFile = lProject.pProjectFile
End Function

Public Sub SaveProject()
Dim i As Integer, msg As String
mdiMain.mnuSaveProject.Caption = "Save Project"
lProject.pProjectSaved = True
If Len(lProject.pProjectFile) = 0 Then
    lProject.pProjectFile = SaveDialog(mdiMain, "NVCD Files (*.nvcd)|*.nvcd|", "Save As ...", App.Path)
    If Len(lProject.pProjectFile) = 0 Then Exit Sub
    lProject.pProjectFile = Left(lProject.pProjectFile, Len(lProject.pProjectFile) - 1) & ".nvcd"
Else
    If DoesFileExist(lProject.pProjectFile) = False Then
        lProject.pProjectFile = SaveDialog(mdiMain, "NVCD Files (*.nvcd)|*.nvcd|", "Save As ...", App.Path)
        If Len(lProject.pProjectFile) = 0 Then Exit Sub
        lProject.pProjectFile = Left(lProject.pProjectFile, Len(lProject.pProjectFile) - 1) & ".nvcd"
    End If
End If
If Len(lProject.pIcon) = 0 Or DoesFileExist(lProject.pIcon) = False Then lProject.pIcon = App.Path & "\DATA\GFX\ICONS\VCD.ICO"
msg = GetFileTitle(lProject.pProjectFile)
msg = Left(lProject.pProjectFile, Len(lProject.pProjectFile) - Len(msg))
WriteINI lProject.pProjectFile, "PROJECT", "NAME", lProject.pName
WriteINI lProject.pProjectFile, "PROJECT", "ICON", lProject.pIcon
WriteINI lProject.pProjectFile, "PROJECT", "MODULES", Trim(Str(lProject.pModuleCount))
WriteINI lProject.pProjectFile, "PROJECT", "FORMS", Trim(Str(lProject.pFormCount))
WriteINI lProject.pProjectFile, "PROJECT", "TYPE", Trim(Str(lProject.pType))
For i = 1 To lProject.pFormCount
    WriteINI lProject.pProjectFile, "FORM" & Trim(Str(i)), "FILE", lProject.pForms(i).fFileName
    WriteINI lProject.pProjectFile, "FORM" & Trim(Str(i)), "NAME", lProject.pForms(i).fName
    WriteINI lProject.pProjectFile, "FORM" & Trim(Str(i)), "CAPTION", lProject.pForms(i).fForm.Caption
    WriteINI lProject.pProjectFile, "FORM" & Trim(Str(i)), "WIDTH", Trim(Str(lProject.pForms(i).fForm.Width))
    WriteINI lProject.pProjectFile, "FORM" & Trim(Str(i)), "HEIGHT", Trim(Str(lProject.pForms(i).fForm.Height))
    WriteINI lProject.pProjectFile, "FORM" & Trim(Str(i)), "LEFT", Trim(Str(lProject.pForms(i).fForm.Left))
    WriteINI lProject.pProjectFile, "FORM" & Trim(Str(i)), "TOP", Trim(Str(lProject.pForms(i).fForm.Top))
    WriteINI lProject.pProjectFile, "FORM" & Trim(Str(i)), "ICON", Trim(Str(lProject.pForms(i).fForm.Icon))
    'SaveFile msg & lProject.pForms(i).fFileName, lProject.pForms(i).fEdit.txtCode.Text
    Stop
    SaveFile msg & lProject.pForms(i).fFileName, lProject.pForms(i).fCode
Next i
For i = 1 To lProject.pModuleCount
    WriteINI lProject.pProjectFile, "MODULE1", "FILE", lProject.pModules(i).mFileName
    WriteINI lProject.pProjectFile, "MODULE1", "NAME", lProject.pModules(i).mName
    'SaveFile msg & lProject.pModules(i).mFileName, lProject.pModules(i).mWindow.txtCode.Text
    SaveFile msg & lProject.pModules(i).mFileName, lProject.pModules(i).mCode
Next i
End Sub

Public Sub SetModuleProperties(lModuleCaption As String)
On Local Error Resume Next
Dim i As Integer, msg As String
'Stop
If Len(lModuleCaption) <> 0 Then
    For i = 0 To lProject.pModuleCount
        If LCase(Trim(lProject.pModules(i).mWindow.Caption)) = Trim(LCase(lModuleCaption)) Then
            If Err.Number = 0 Then
                With frmProporties.ctlPropertyList1
                    .ClearPropList
                    .AddProporty "(Name)", lProject.pModules(i).mName, nText
                    .AddProporty "FileName", lProject.pModules(i).mFileName, nText
                    
                    Exit For
                End With
            Else
                If Err.Number = 91 Then Err.Clear
            End If
        End If
    Next i
End If
End Sub

Public Sub SetButtonProperties(lFormIndex As Integer, lButtonIndex As Integer)
On Local Error Resume Next
Dim i As Integer, msg As String
With frmProporties.ctlPropertyList1
    'msg = "Black" & vbCrLf & "White" & vbCrLf & "Red" & vbCrLf & "Blue" & vbCrLf & "Green" & vbCrLf & "Yellow" & vbCrLf & "Orange" & vbCrLf & "Cyan"
    .ClearPropList
    '.AddProporty "(Name)", lProject.pForms(i).fName, nText
    .AddProporty "(Name)", lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex).Tag, nText
    .AddProporty "Caption", lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex).Caption, nText
    .AddProporty "Left", lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex).Left, nText
    .AddProporty "Top", lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex).Top, nText
    .AddProporty "Width", lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex).Width, nText
    .AddProporty "Height", lProject.pForms(lFormIndex).fForm.cmdButton(lButtonIndex).Height, nText
    '.AddProporty "BackColor", lProject.pForms(i).fForm.BackColor, nDropDown, msg
End With
End Sub

Public Sub SetFormBackColor(lIndex As Integer, lBackColor As String)
'Stop
Select Case LCase(lBackColor)
Case "blue"
    lProject.pForms(lIndex).fForm.BackColor = vbBlue
End Select
End Sub

Public Sub SetFormCaption(lIndex As Integer, lName As String)
If lIndex <> 0 Then
    lProject.pForms(lIndex).fForm.Caption = lName
Else
    MsgBox "FATAL ERROR"
End If
End Sub

Public Sub SetFormName(lIndex As Integer, lName As String)
On Local Error Resume Next
With lProject.pForms(lIndex)
    .fName = lName
    .fForm.Tag = lName
End With
End Sub

Public Sub SetFormProperties(lFormIndex As Integer)
On Local Error Resume Next
Dim i As Integer, msg As String
With frmProporties.ctlPropertyList1
    msg = "Black" & vbCrLf & "White" & vbCrLf & "Red" & vbCrLf & "Blue" & vbCrLf & "Green" & vbCrLf & "Yellow" & vbCrLf & "Orange" & vbCrLf & "Cyan"
    .ClearPropList
    .AddProporty "(Name)", lProject.pForms(lFormIndex).fName, nText
    .AddProporty "BackColor", lProject.pForms(lFormIndex).fForm.BackColor, nDropDown, msg
    .AddProporty "Caption", lProject.pForms(lFormIndex).fForm.Caption, nText
    
End With
End Sub

Public Function AddModule(lName As String, lFile As String, Optional lShow As Boolean, Optional lAddSubMain As Boolean) As Integer
Dim lNode As Node
If Len(lName) <> 0 Then
    lProject.pModuleCount = lProject.pModuleCount + 1
    AddModule = lProject.pModuleCount
    With lProject.pModules(lProject.pModuleCount)
        lProject.pProjectSaved = False
        .mFileName = lFile
        .mName = lName
        Set .mWindow = New frmEdit
        .mWindow.SetModIndex lProject.pModuleCount
        If lShow = True Then
            .mWindow.Caption = lFile
            .mWindow.Show
            If lAddSubMain = True Then
                .mWindow.txtCode.Text = "Sub Main()" & vbCrLf & "" & vbCrLf & "End Sub" & vbCrLf
            End If
        End If
        If lAddSubMain = True Then
            .mCode = "Sub Main()" & vbCrLf & "" & vbCrLf & "End Sub" & vbCrLf
        End If
        Set lNode = frmProporties.tvwProject.Nodes.Add(2, tvwChild, , GetFileTitle(lFile), 1)
        lNode.Tag = Trim(Str(lProject.pModuleCount))
        frmProporties.tvwProject.Nodes(2).Expanded = True
    End With
End If
End Function

Public Sub SetModuleCursor(lModIndex As Integer, lPosition As Integer)
lProject.pModules(lModIndex).mEditCursor = lPosition
End Sub

Public Sub ShowModule(lIndex As Integer)
If lIndex <> 0 Then
    lProject.pModules(lIndex).mWindow.Caption = lProject.pModules(lIndex).mFileName
    lProject.pModules(lIndex).mWindow.Show
    lProject.pModules(lIndex).mWindow.txtCode.SelStart = lProject.pModules(lIndex).mEditCursor
End If
End Sub

Public Sub SetModuleFocus(lIndex As Integer, lVisible As Boolean)
lProject.pModules(lIndex).mWindow.Visible = lVisible
Select Case lVisible
Case True
    lProject.pModules(lIndex).mWindow.Caption = GetFileTitle(lProject.pModules(lIndex).mFileName)
    lProject.pModules(lIndex).mWindow.txtCode.Text = lProject.pModules(lIndex).mCode
    lProject.pModules(lIndex).mWindow.SetModIndex lIndex
Case False
End Select
lProject.pModules(lIndex).mWindow.SetFocus
End Sub

Public Sub SetFormFocus(lIndex As Integer)
With lProject.pForms(lIndex)
    .fForm.Visible = True
    .fForm.SetFocus
    .fForm.Caption = GetFileTitle(.fFileName)
    .fForm.SetFormIndex lIndex
End With
End Sub

Public Function ReturnProjectName() As String
ReturnProjectName = lProject.pName
End Function

Public Sub ClearProject()
Dim i As Integer
With lProject
    For i = 1 To 128
        Set .pForms(i).fEdit = Nothing
        .pForms(i).fFileName = ""
        Set .pForms(i).fForm = Nothing
        .pForms(i).fName = ""
    Next i
    .pFormCount = 0
    .pIcon = ""
    For i = 1 To 128
        .pModules(i).mFileName = ""
        .pModules(i).mName = ""
        Set .pModules(i).mWindow = Nothing
    Next i
    .pModuleCount = 0
    .pName = ""
    .pProjectFile = ""
    .pProjectSaved = False
'    .pType = 0
End With
mdiMain.Caption = App.Title
frmProporties.FormDragger1.Caption = "Project"
End Sub

Public Sub SetProjectName(lData As String)
lProject.pProjectSaved = False
lProject.pName = lData
End Sub

Public Sub SetFormCode(lIndex As Integer, lCode As String)
If lIndex <> 0 Then
    With lProject.pForms(lIndex)
        'MsgBox "Setting form Code: " & vbCrLf & lCode
        .fCode = lCode
    End With
End If
End Sub

Public Function ReturnProjectSaved() As Boolean
ReturnProjectSaved = lProject.pProjectSaved
End Function

Public Sub ShowFormCode(lIndex As Integer)
'Stop
If lIndex <> 0 Then
    With lProject.pForms(lIndex)
        If Len(.fCode) = 0 Then
            .fCode = "Sub Form_Load()" & vbCrLf & "" & vbCrLf & "End Sub" & vbCrLf
        End If
        .fEdit.txtCode.Text = .fCode
        .fEdit.txtCode.SelStart = .fEditCursor
        .fEdit.Show
        .fEdit.Visible = True
        .fEdit.SetFormIndex lIndex
    End With
End If
End Sub

Public Sub SetFormCodeEdit(lIndex As Integer, lCode As String)
lProject.pForms(lIndex).fEdit.txtCode.Text = lCode
End Sub

Public Sub ShowFormCodeByIndex(lIndex As Integer)
Dim l As Integer
With lProject.pForms(lIndex)
'    .fEdit.Visible = True
'    .fEdit.txtCode.Text = .fCode
'    .fEdit.Caption = GetFileTitle(.fFileName)
    If Len(.fCode) = 0 Then
        .fCode = "Sub Form_Load()" & vbCrLf & "" & vbCrLf & "End Sub" & vbCrLf
        l = 17
    End If
    .fEdit.txtCode.Text = .fCode
    .fEdit.Show
    .fEdit.Visible = True
    .fEdit.SetFormIndex lIndex
    .fEdit.Caption = GetFileTitle(.fFileName)
    .fEdit.txtCode.SelStart = .fEditCursor
    .fEdit.ApplyCodeInsert cForm
    .fEdit.SetTextFocus "Form_Load"
End With
End Sub

Public Function ReturnFormCount() As Integer
ReturnFormCount = lProject.pFormCount
End Function

Public Function ReturnFormName(lIndex As Integer) As String
ReturnFormName = lProject.pForms(lIndex).fName
End Function

Public Sub SetModuleCode(lIndex As Integer, lCode As String)
If lIndex <> 0 And Len(lCode) <> 0 Then lProject.pModules(lIndex).mCode = lCode
End Sub

Public Function DoesSubExist(lFormIndex As Integer, lSub As String) As Boolean
Dim msg() As String, i As Integer, msg2 As String, b As Boolean
msg = Split(lProject.pForms(lFormIndex).fCode, vbCrLf)
For i = 0 To UBound(msg)
    If LCase(Trim(msg(i))) = LCase(Trim(lSub)) Then
        DoesSubExist = True
        Exit For
    End If
Next i
End Function

Public Function ReturnFormSub(lFormIndex As Integer, lSub As String) As String
Dim msg() As String, i As Integer, msg2 As String, b As Boolean
msg = Split(lProject.pForms(lFormIndex).fCode, vbCrLf)
For i = 0 To UBound(msg)
    If LCase(Trim(msg(i))) = LCase(Trim(lSub)) Then b = True
    If b = True Then
        If Len(msg2) <> 0 Then
            msg2 = msg2 & vbCrLf & msg(i)
        Else
            msg2 = msg(i)
        End If
    End If
    If Trim(LCase(msg(i))) = "end sub" Then
        If Len(msg2) <> 0 Then
            b = False
            ReturnFormSub = msg2
            Exit For
        End If
    End If
Next i
End Function

Public Function AddObjectToForm(lIndex As Integer, lData As String)
Dim msg As String, msg2() As String, msg3 As String, i As Integer
If DoesSubExist(lIndex, "Sub Form_Objects()") = True Then
    msg = ReturnFormSub(lIndex, "Sub Form_Objects()")
    If Len(msg) <> 0 Then
        msg2 = Split(msg, vbCrLf)
        For i = 0 To UBound(msg2)
            If Trim(LCase(msg2(i))) = "end sub" Then
                msg3 = msg3 & vbCrLf & lData & vbCrLf & "End Sub"
            Else
                msg3 = msg3 & vbCrLf & msg2(i)
            End If
        Next i
        lProject.pForms(lIndex).fCode = msg3
        lProject.pForms(lIndex).fEdit.txtCode.Text = msg3
    End If
Else
    lProject.pForms(lIndex).fCode = "Sub Form_Objects()" & vbCrLf & lData & vbCrLf & "End Sub" & vbCrLf & lProject.pForms(lIndex).fCode
End If
'lProject.pForms(lIndex).fCode = lData & vbCrLf & lProject.pForms(lIndex).fCode
End Function

Public Function AddForm(lName As String) As Integer
Dim i As Integer, b As Boolean, lNode As Node
If Len(lProject.pName) <> 0 Then
    lProject.pProjectSaved = False
    lProject.pFormCount = lProject.pFormCount + 1
    AddForm = lProject.pFormCount
    With lProject.pForms(lProject.pFormCount)
        .fName = lName
        .fFileName = lName & ".txt"
        Set .fForm = New frmForm
        Set .fEdit = New frmEdit
        .fForm.Visible = True
        .fForm.Caption = lName
        .fEdit.Caption = lName & ".txt"
        .fEdit.txtCode.Text = "Sub Form_Load()" & vbCrLf & vbCrLf & "End Sub" & vbCrLf
        .fEdit.Visible = False
        .fForm.SetFormIndex lProject.pFormCount
        For i = 1 To frmProporties.tvwProject.Nodes.Count
            If frmProporties.tvwProject.Nodes(i).Text = "Forms" Then
                b = True
                Exit For
            End If
        Next i
        Set lNode = frmProporties.tvwProject.Nodes.Add(3, tvwChild, , lName & ".txt", 1)
        lNode.Tag = lProject.pFormCount
        frmProporties.tvwProject.Nodes(3).Expanded = True
    End With
End If
End Function

