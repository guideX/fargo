Attribute VB_Name = "mdlMain"
Option Explicit
Private lEXEPassword As String
Private lExecute As clsExecute
Private lBag As New PropertyBag
Private lErrHandler As New clsErrHandler
Private lPropBagOpened As Boolean

Public Sub SetFormLeft(lIndex As Integer, lLeft As Integer)
On Local Error GoTo ErrHandler
lExecute.SetFormLeft lIndex, lLeft
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormLeft(lIndex As Integer, lLeft As Integer)"
End Sub

Public Sub SetFormTop(lIndex As Integer, lTop As Integer)
On Local Error GoTo ErrHandler
lExecute.SetFormTop lIndex, lTop
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormTop(lIndex As Integer, lLeft As Integer)"
End Sub

Public Sub SetFormHeight(lIndex As Integer, lHeight As Integer)
On Local Error GoTo ErrHandler
lExecute.SetFormHeight lIndex, lHeight
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormHeight(lIndex As Integer, lHeight As Integer)"
End Sub

Public Sub SetFormWidth(lIndex As Integer, lWidth As Integer)
On Local Error GoTo ErrHandler
lExecute.SetFormHeight lIndex, lWidth
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormWidth(lIndex As Integer, lWidth As Integer)"
End Sub

Public Function ReturnFormName(lIndex As Integer)
On Local Error GoTo ErrHandler
ReturnFormName = lExecute.ReturnFormName(lIndex)
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnEXEPassword() As String"
End Function

Public Function ReturnEXEPassword() As String
On Local Error GoTo ErrHandler
ReturnEXEPassword = lEXEPassword
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnEXEPassword() As String"
End Function

Public Sub Trigger_cmdButton_GotFocus(lFormIndex As Integer, lIndex As Integer, lName As String)
On Local Error GoTo ErrHandler
lExecute.ProcessSub lExecute.ReturnFormSub(lFormIndex, "Sub " & lName & "_GotFocus()")
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_cmdButton_GotFocus(lFormIndex As Integer, lIndex As Integer, lName As String)"
End Sub

Public Sub Trigger_cmdButton_LostFocus(lFormIndex As Integer, lIndex As Integer, lName As String)
On Local Error GoTo ErrHandler
lExecute.ProcessSub lExecute.ReturnFormSub(lFormIndex, "Sub " & lName & "_LostFocus()")
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_cmdButton_LostFocus(lFormIndex As Integer, lIndex As Integer, lName As String)"
End Sub

Public Sub Trigger_cmdButton_Click(lFormIndex As Integer, lIndex As Integer, lName As String)
On Local Error GoTo ErrHandler
lExecute.ProcessSub lExecute.ReturnFormSub(lFormIndex, "Sub " & lName & "_Click()")
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_cmdButton_Click(lIndex As Integer, lName As String)"
End Sub

Public Sub Trigger_Form_Position(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_Position lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_Click(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_Activate(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_Activate lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_Click(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_Click(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_Click lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_Click(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_Unload(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_Unload lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_Resize(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_Load(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_Load lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_Unload(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_Objects(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_Objects lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_Objects(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_LostFocus(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_LostFocus lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_LostFocus(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_Resize(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_Resize lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_Resize(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_GotFocus(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_GotFocus lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_GotFocus(lIndex As Integer)"
End Sub

Public Sub Trigger_Form_DblClick(lIndex As Integer)
On Local Error GoTo ErrHandler
lExecute.Event_Form_DblClick lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Trigger_Form_DblClick(lIndex As Integer)"
End Sub

Public Sub ProcessRuntimeError(lNumber As Long, lDescription As String, lSub As String)
On Local Error GoTo ErrHandler
lErrHandler.ProcessError lNumber, lDescription, lSub
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub ProcessRuntimeError(lNumber As Long, lDescription As String, lSub As String)"
End Sub

Sub Main()
On Local Error GoTo ErrHandler
Dim mbox As VbMsgBoxResult, msg As String, msg2 As String, lForm As Form, lCommand As String, lRun As Boolean, i As Integer
lEXEPassword = "n3h4Vm3I92F"
Set lErrHandler = New clsErrHandler
Set lExecute = New clsExecute
If KeyQuality(lEXEPassword) < 20 Then End
LoadPaths
lCommand = Command$
If Len(lCommand) <> 0 Then
    If InStr(lCommand, " -run") Then
        lCommand = Replace(lCommand, " -run", "")
        lRun = True
    End If
    If lExecute.LoadProjectFromFile(lCommand) = False Then
        MsgBox "Project file not found", vbExclamation
        End
    End If
    If lRun = False Then
        mbox = MsgBox("Would you like to compile the following project: " & vbCrLf & lCommand & "?", vbYesNoCancel + vbQuestion)
    Else
        mbox = vbNo
    End If
    Select Case mbox
    Case vbYes
SaveDLG:
        RunProject
        DoEvents
        msg = SaveDialog(frmNothing, "EXE Files (*.exe)|*.exe|", "Compile EXE", App.Path)
        If Len(msg) <> 0 Then
            msg = Left(msg, Len(msg) - 1) & ".exe"
            msg2 = GetFileTitle(msg)
            msg = Left(msg, Len(msg) - Len(msg2))
            If DoesFileExist(msg & msg2) = True Then
                MsgBox "File Exists, please select another", vbCritical
                msg = ""
                msg2 = ""
                GoTo SaveDLG
            End If
            lExecute.CompileProject msg, msg2, True, True
            MsgBox msg2 & " Compiled", vbInformation
            End
        Else
            End
        End If
    Case vbNo
        RunProject
    Case vbCancel
        End
    End Select
Else
    If lPropBagOpened = False Then
        lPropBagOpened = True
        OpenPropBag
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Sub Main()"
End Sub

Private Sub lExecute_RuntimeError(lRuntimeError As eRuntimeError)

End Sub

Public Sub ReachCompiler(lFilePath As String, lFileName As String)
On Local Error GoTo ErrHandler
lExecute.CompileProject lFilePath, lFileName
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub ReachCompiler(lFilePath As String, lFileName As String)"
End Sub

Public Function ReturnProj() As String
On Local Error GoTo ErrHandler
ReturnProj = lExecute.ReturnProjectName
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnProj() As String"
End Function

Public Sub SetProjectName(lName As String)
On Local Error GoTo ErrHandler
lExecute.SetProjectName lName
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetProjectName(lName As String)"
End Sub

Public Sub ManSetModCount(lCount As Integer)
On Local Error GoTo ErrHandler
lExecute.SetModCount lCount
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub ManSetModCount(lCount As Integer)"
End Sub

Public Sub ManSetMod(lIndex As Integer, lCode As String, lName As String)
On Local Error GoTo ErrHandler
lExecute.SetMod lIndex, lName, lCode
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub ManSetMod(lIndex As Integer, lCode As String, lName As String)"
End Sub

Public Sub RunProject()
On Local Error GoTo ErrHandler
lExecute.RunMain
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub RunProject()"
End Sub

Public Sub ManSetSubMain(lData As String)
On Local Error GoTo ErrHandler
lExecute.SetSubMain lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub ManSetSubMain(lData As String)"
End Sub

Public Sub OpenPropBag()
On Local Error GoTo ErrHandler
Dim l As Long, v As Variant, b() As Byte, i As Integer, f As Integer, msg As String, mbox As VbMsgBoxResult
If KeyQuality(ReturnEXEPassword) < 20 Then
    MsgBox "Key does not meet minimum requirements !", vbCritical
    End
End If
Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
    Get #1, LOF(1) - 3, l
    Seek #1, l
    Get #1, , v
    b = v
    lBag.Contents = b
    lBag.WriteProperty "LOF", LOF(1)
    lBag.WriteProperty "l", l
Close #1
With lBag
    lExecute.SetSubMain DecodeString(DecodeStr64(.ReadProperty("SUBMAIN")), ReturnEXEPassword, True)
    lExecute.SetProjectName DecodeString(DecodeStr64(.ReadProperty("PROJECTNAME")), ReturnEXEPassword, True)
    lExecute.SetIcon DecodeString(DecodeStr64(.ReadProperty("PROJECTICON")), ReturnEXEPassword, True)
    msg = DecodeString(DecodeStr64(.ReadProperty("PROJECTDATA")), ReturnEXEPassword, True)
    f = Int(.ReadProperty("FORMCOUNT"))
    lExecute.SetFormCount f
    For i = 1 To f
        If Len(.ReadProperty("FORMNAME" & Trim(Str(i)), "")) <> 0 Then
            lExecute.SetForm i
            lExecute.SetFormCode i, DecodeString(DecodeStr64(.ReadProperty("FORMCODE" & Trim(Str(i)), "")), ReturnEXEPassword, True)
            lExecute.SetFormName i, DecodeString(DecodeStr64(.ReadProperty("FORMNAME" & Trim(Str(i)), "")), ReturnEXEPassword, True)
            lExecute.SetFormCaption i, DecodeString(DecodeStr64(.ReadProperty("FORMCAPTION" & Trim(Str(i)), "")), ReturnEXEPassword, True)
            lExecute.SetFormPosition i, DecodeString(DecodeStr64(.ReadProperty("FORMLEFT" & Trim(Str(i)), 0)), ReturnEXEPassword, True), DecodeString(DecodeStr64(.ReadProperty("FORMTOP" & Trim(Str(i)), 0)), ReturnEXEPassword, True), DecodeString(DecodeStr64(.ReadProperty("FORMWIDTH" & Trim(Str(i)), 0)), ReturnEXEPassword, True), DecodeString(DecodeStr64(.ReadProperty("FORMHEIGHT" & Trim(Str(i)), 0)), ReturnEXEPassword, True)
            lExecute.SetFormSubs lExecute.ReturnFormCode(i), i
        End If
    Next i
    f = Int(.ReadProperty("MODCOUNT"))
    For i = 1 To f
        lExecute.SetMod i, DecodeString(DecodeStr64(.ReadProperty("MODNAME" & Trim(Str(i)), "")), ReturnEXEPassword, True), DecodeString(DecodeStr64(.ReadProperty("MODDATA" & Trim(Str(i)), "")), ReturnEXEPassword, True)
    Next i
End With
If Len(lExecute.ReturnProjectName) = 0 Then
    End
Else
    lExecute.RunMain
    'End
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub OpenPropBag()"
End Sub
