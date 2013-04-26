VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Enum eObjectTypes
    oUnknown = 0
    oCommandButton = 1
End Enum
Enum eRuntimeError
    rUnknown = 0
    rProjectFileMissing = 1
    rProjectFileWrongPath = 2
    rTooManyParameters = 3
End Enum
Private Type gFormSubs
    sForm_LostFocus As String
    sForm_GotFocus As String
    sForm_Load As String
    sForm_Click As String
    sForm_DblClick As String
    sForm_Unload As String
    sForm_Resize As String
    sForm_Objects As String
    sForm_Activate As String
    sForm_Subs As String
    sForm_Position As String
End Type
Private Type gForm
    fSubs As gFormSubs
    fForm As frmNothing
    fName As String
    fCode As String
End Type
Private Type gModule
    mName As String
    mCode As String
End Type
Private Type gObjectSubs
    oObject_Click As String
    oObject_GotFocus As String
End Type
Private Type gObject
    oName As String
    oType As eObjectTypes
    oSubs As gObjectSubs
End Type
Private Type gProject
    pIconPath As String
    pFileName As String
    pName As String
    pFormCount As Integer
    pModuleCount As Integer
    pForms(128) As gForm
    pModules(128) As gModule
    pSubMain As String
End Type
Private lProject As gProject
Private lCurrentForm As Integer
Public Event RuntimeError(lRuntimeError As eRuntimeError)

Public Function ReturnFormName(lFormIndex As Integer) As String
On Local Error GoTo ErrHandler
ReturnFormName = lProject.pForms(lFormIndex).fName
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnFormName(lFormIndex As Integer) As String"
End Function

Public Function ReturnFormCode(lFormIndex As Integer) As String
On Local Error GoTo ErrHandler
ReturnFormCode = lProject.pForms(lFormIndex).fCode
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub ReturnFormCode(lFormIndex as Integer) As String"
End Function

Public Sub ChangeRuntimeDLL(lEXEToPatch As String, lOldDLL As String, lNewDLL As String)
On Local Error GoTo ErrHandler
Dim c As Boolean, msg As String * 1024, l As Long, i As Integer, o As Long, n As Long, msg2 As String, g As Long, l2 As Long
If lNewDLL = "" Or c = True Or lOldDLL = "" Or lEXEToPatch = "" Or FileLen(lEXEToPatch) <= 0 Or Len(lOldDLL) < Len(lNewDLL) Then c = True
If c = True Then Exit Sub
l = Len(msg)
msg2 = lNewDLL + String(Len(lOldDLL) - Len(lNewDLL), vbNullChar)
n = Len(lOldDLL)
o = FileLen(lEXEToPatch)
g = 1
i = FreeFile
Open lEXEToPatch For Binary As i
    While Not EOF(i)
        Get i, g, msg
        l2 = InStr(msg, lOldDLL)
        If l2 > 0 Then
            Put i, g + l2 - 1, msg2
        End If
        g = g + l - n
    Wend
Close i
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub ChangeRuntimeDLL(lEXEToPatch As String, lOldDLL As String, lNewDLL As String)"
End Sub

Public Sub SetIcon(lPath As String)
On Local Error GoTo ErrHandler
lProject.pIconPath = lPath
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetIcon(lPath As String)"
End Sub

Public Sub SetSubMain(lData As String)
On Local Error GoTo ErrHandler
lProject.pSubMain = lData
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetSubMain(lData As String)"
End Sub

Public Sub RunMain()
On Local Error GoTo ErrHandler
ProcessSub lProject.pSubMain
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub RunMain()"
End Sub

Public Sub SetProjectName(lName As String)
On Local Error GoTo ErrHandler
lProject.pName = lName
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetProjectName(lName As String)"
End Sub

Public Sub SetMod(lIndex As Integer, lName As String, lCode As String)
On Local Error GoTo ErrHandler
lProject.pModules(lIndex).mName = lName
lProject.pModules(lIndex).mCode = lCode
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetMod(lIndex As Integer, lName As String, lCode As String)"
End Sub

Public Sub SetModCount(lCount As Integer)
On Local Error GoTo ErrHandler
lProject.pModuleCount = lCount
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetModCount(lCount As Integer)"
End Sub

Public Function ReturnProjectName() As String
On Local Error GoTo ErrHandler
ReturnProjectName = lProject.pName
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnProjectName() As String"
End Function

Public Sub ShowForm(lIndex As Integer)
On Local Error GoTo ErrHandler
lProject.pForms(lIndex).fForm.Show
lProject.pForms(lIndex).fForm.Visible = True
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub ShowForm(lIndex As Integer)"
End Sub

Public Sub SetForm(lIndex As Integer)
On Local Error GoTo ErrHandler
Set lProject.pForms(lIndex).fForm = New frmNothing
lProject.pForms(lIndex).fForm.SetFormIndex lIndex
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetForm(lIndex As Integer)"
End Sub

Public Sub SetFormPosition(lIndex As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer)
On Local Error GoTo ErrHandler
lProject.pForms(lIndex).fForm.Left = lLeft
lProject.pForms(lIndex).fForm.Top = lTop
lProject.pForms(lIndex).fForm.Width = lWidth
lProject.pForms(lIndex).fForm.Height = lHeight
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormPosition(lIndex As Integer, lLeft As Integer, lTop As Integer, lWidth As Integer, lHeight As Integer)"
End Sub

Public Sub SetFormName(lIndex As Integer, lName As String)
On Local Error GoTo ErrHandler
lProject.pForms(lIndex).fName = lName
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormName(lIndex As Integer, lName As String)"
End Sub

Public Sub SetFormCaption(lIndex As Integer, lCaption As String)
On Local Error GoTo ErrHandler
lProject.pForms(lIndex).fForm.Caption = lCaption
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormCaption(lIndex As Integer, lCaption As String)"
End Sub

Public Sub SetFormCount(lCount As Integer)
On Local Error GoTo ErrHandler
lProject.pFormCount = lCount
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormCount(lCount As Integer)"
End Sub

Public Sub SetFormCode(lIndex As Integer, lCode As String)
On Local Error GoTo ErrHandler
lProject.pForms(lIndex).fCode = lCode
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub SetFormCount(lCount As Integer)"
End Sub

Public Sub CompileProject(lOutputPath As String, lOutputFile As String, Optional lChangeDLL As Boolean, Optional lChangeIcon As Boolean)
On Local Error GoTo ErrHandler
Dim BeginPos As Long, lBag As New PropertyBag, varTemp As Variant, lerr As String, i As Integer
If KeyQuality(ReturnEXEPassword) < 20 Then
    End
End If
With lBag
    .WriteProperty "FORMCOUNT", lProject.pFormCount
    For i = 1 To lProject.pFormCount
        .WriteProperty "FORMCODE" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pForms(i).fCode, ReturnEXEPassword, True), 68)
        .WriteProperty "FORMNAME" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pForms(i).fName, ReturnEXEPassword, True), 68)
        .WriteProperty "FORMCAPTION" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pForms(i).fForm.Caption, ReturnEXEPassword, True), 68)
        .WriteProperty "FORMWIDTH" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pForms(i).fForm.Width, ReturnEXEPassword, True), 68)
        .WriteProperty "FORMHEIGHT" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pForms(i).fForm.Height, ReturnEXEPassword, True), 68)
        .WriteProperty "FORMTOP" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pForms(i).fForm.Top, ReturnEXEPassword, True), 68)
        .WriteProperty "FORMLEFT" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pForms(i).fForm.Left, ReturnEXEPassword, True), 68)
    Next i
    .WriteProperty "MODCOUNT", lProject.pModuleCount
    For i = 0 To lProject.pModuleCount
        .WriteProperty "MODNAME" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pModules(i).mName, ReturnEXEPassword, True), 68)
        .WriteProperty "MODDATA" & Trim(Str(i)), EncodeStr64(EncodeString(lProject.pModules(i).mCode, ReturnEXEPassword, True), 68)
    Next i
    .WriteProperty "PROJECTNAME", EncodeStr64(EncodeString(lProject.pName, ReturnEXEPassword, True), 68)
    .WriteProperty "PROJECTICON", EncodeStr64(EncodeString(lProject.pIconPath, ReturnEXEPassword, True), 68)
    .WriteProperty "SUBMAIN", EncodeStr64(EncodeString(lProject.pSubMain, ReturnEXEPassword, True), 68)
    .WriteProperty "PROJECTDATA", EncodeStr64(EncodeString(ReadFile(lProject.pFileName), ReturnEXEPassword, True), 68)
End With
FileCopy App.Path & "\NVCDI.EXE", lOutputPath & "\" & lOutputFile
If lChangeDLL = True Then ChangeRuntimeDLL lOutputPath & lOutputFile, "MSVBVM60.DLL", "nvcdi.dll"
If lChangeIcon = True Then
    If DoesFileExist(lProject.pIconPath) = True Then ReplaceIcons lProject.pIconPath, lOutputPath & lOutputFile, lerr
End If
DoEvents
If Right(lOutputPath, 1) <> "\" Then lOutputPath = lOutputPath & "\"
Open lOutputPath & lOutputFile For Binary As #1
    BeginPos = LOF(1)
    varTemp = lBag.Contents
    Seek #1, LOF(1)
    Put #1, , varTemp
    Put #1, , BeginPos
Close #1
DoEvents
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub CompileProject(lOutputPath As String, lOutputFile As String, Optional lChangeDLL As Boolean, Optional lChangeIcon As Boolean)"
End Sub

Public Function ReturnProjectPath() As String
On Local Error GoTo ErrHandler
ReturnProjectPath = Left(lProject.pFileName, Len(lProject.pFileName) - Len(GetFileTitle(lProject.pFileName)))
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnProjectPath() As String"
End Function

Public Sub SetFormSubs(lCode As String, lFormIndex As Integer)
If InStr(LCase(lCode), "sub form_objects()") Then lProject.pForms(lFormIndex).fSubs.sForm_Objects = ReturnFormSub(lFormIndex, "Sub Form_Objects()")
If InStr(LCase(lCode), "Sub form_subs()") Then
    lProject.pForms(lFormIndex).fSubs.sForm_Subs = ReturnFormSub(lFormIndex, "Sub Form_Subs()")
End If
If InStr(LCase(lCode), "sub form_lostfocus()") Then lProject.pForms(lFormIndex).fSubs.sForm_LostFocus = ReturnFormSub(lFormIndex, "Sub Form_LostFocus()")
If InStr(LCase(lCode), "sub form_load()") Then lProject.pForms(lFormIndex).fSubs.sForm_Load = ReturnFormSub(lFormIndex, "Sub Form_Load()")
If InStr(LCase(lCode), "sub form_click()") Then lProject.pForms(lFormIndex).fSubs.sForm_Click = ReturnFormSub(lFormIndex, "Sub Form_Click()")
If InStr(LCase(lCode), "sub form_dblclick()") Then lProject.pForms(lFormIndex).fSubs.sForm_DblClick = ReturnFormSub(lFormIndex, "Sub Form_DblClick()")
If InStr(LCase(lCode), "sub form_unload()") Then lProject.pForms(lFormIndex).fSubs.sForm_Unload = ReturnFormSub(lFormIndex, "Sub Form_Unload()")
If InStr(LCase(lCode), "sub form_gotfocus()") Then lProject.pForms(lFormIndex).fSubs.sForm_GotFocus = ReturnFormSub(lFormIndex, "Sub Form_GotFocus()")
If InStr(LCase(lCode), "sub form_resize()") Then lProject.pForms(lFormIndex).fSubs.sForm_Resize = ReturnFormSub(lFormIndex, "Sub Form_Resize()")
If InStr(LCase(lCode), "sub form_position()") Then lProject.pForms(lFormIndex).fSubs.sForm_Position = ReturnFormSub(lFormIndex, "Sub Form_Position()")
End Sub

Public Function LoadProjectFromFile(lFile As String) As Boolean
On Local Error GoTo ErrHandler
Dim msg As String, i As Integer, msg2 As String, msg3 As String, msg4 As String, l As Integer
If Len(lFile) <> 0 Then
    If DoesFileExist(lFile) = False Then
        If Not InStr(lFile, "\") Then
            For i = 1 To ReturnPathCount
                msg2 = FindFilePath(ReturnPath(i), lFile)
                If Len(msg2) <> 0 Then
                    lFile = msg2
                    If DoesFileExist(lFile) = True Then
                        Exit For
                    End If
                End If
            Next i
        Else
            RaiseEvent RuntimeError(rProjectFileWrongPath)
            Exit Function
        End If
        If DoesFileExist(lFile) = False Then
            Exit Function
        End If
    End If
    If Len(lFile) = 0 Then
        RaiseEvent RuntimeError(rProjectFileWrongPath)
        Exit Function
    Else
        lProject.pFileName = lFile
    End If
    LoadProjectFromFile = True
    lProject.pFormCount = Int(ReadINI(lFile, "PROJECT", "FORMS", 0))
    For i = 1 To lProject.pFormCount
        With lProject.pForms(i)
            Set .fForm = New frmNothing
            .fName = ReadINI(lFile, "FORM" & Trim(Str(i)), "NAME", "")
            If Len(.fName) <> 0 Then
                .fForm.Caption = ReadINI(lFile, "FORM" & Trim(Str(i)), "CAPTION", "")
                .fForm.Left = Int(ReadINI(lFile, "FORM" & Trim(Str(i)), "LEFT", 0))
                .fForm.Top = Int(ReadINI(lFile, "FORM" & Trim(Str(i)), "TOP", 0))
                .fForm.Width = Int(ReadINI(lFile, "FORM" & Trim(Str(i)), "WIDTH", 2500))
                .fForm.Height = Int(ReadINI(lFile, "FORM" & Trim(Str(i)), "HEIGHT", 2000))
                msg3 = ReadINI(Left(lFile, Len(lFile)), "FORM" & Trim(Str(i)), "FILE", "")
                If DoesFileExist(msg3) = False Then
                    msg4 = FindFilePath(ReturnFilePath(lFile), msg3)
                    If DoesFileExist(msg4) = True Then
                        .fCode = ReadFile(msg4)
                    End If
                Else
                    .fCode = ReadFile(msg3)
                End If
                SetFormSubs .fCode, i
            End If
        End With
    Next i
    lProject.pName = ReadINI(lFile, "PROJECT", "NAME", "")
    lProject.pIconPath = ReadINI(lFile, "PROJECT", "ICON", "")
    If DoesFileExist(lProject.pIconPath) = False Then
        If DoesFileExist(Left(lProject.pFileName, Len(lProject.pFileName) - Len(GetFileTitle(lProject.pFileName))) & lProject.pIconPath) = True Then
            lProject.pIconPath = Left(lProject.pFileName, Len(lProject.pFileName) - Len(GetFileTitle(lProject.pFileName))) & lProject.pIconPath
        Else
            MsgBox "Icon file not found"
        End If
    End If
    Select Case Len(lProject.pName)
    Case 0
        RaiseEvent RuntimeError(rProjectFileMissing)
    Case Else
        lProject.pModuleCount = Int(ReadINI(lFile, "PROJECT", "MODULES", 0))
    End Select
    For i = 1 To lProject.pModuleCount
        msg3 = ReadINI(Left(lFile, Len(lFile)), "MODULE" & Trim(Str(i)), "FILE", "")
        If DoesFileExist(msg3) = False Then
            msg4 = FindFilePath(ReturnFilePath(lFile), msg3)
            If DoesFileExist(msg4) = True Then
                lProject.pModules(i).mCode = ReadFile(msg4)
            End If
        Else
            lProject.pModules(i).mCode = ReadFile(msg3)
        End If
        If InStr(LCase(lProject.pModules(i).mCode), "sub main()") Then lProject.pSubMain = ReturnSub(i, "Sub Main()")
    Next i
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function LoadProjectFromFile(lFile As String) As Boolean"
End Function

Public Function ReturnSub(lModuleIndex As Integer, lSub As String) As String
On Local Error GoTo ErrHandler
Dim msg() As String, i As Integer, msg2 As String, b As Boolean
msg = Split(lProject.pModules(lModuleIndex).mCode, vbCrLf)
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
        b = False
        ReturnSub = msg2
        Exit For
    End If
Next i
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnSub(lModuleIndex As Integer, lSub As String) As String"
End Function

Public Function ReturnFormSub(lFormIndex As Integer, lSub As String) As String
On Local Error GoTo ErrHandler
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
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnFormSub(lFormIndex As Integer, lSub As String) As String"
End Function

Public Sub Event_Form_Unload(lIndex As Integer)
On Local Error GoTo ErrHandler
ProcessSub lProject.pForms(lIndex).fSubs.sForm_Unload
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_Unload(lIndex As Integer)"
End Sub

Public Sub Event_Form_GotFocus(lIndex As Integer)
On Local Error GoTo ErrHandler
ProcessSub lProject.pForms(lIndex).fSubs.sForm_GotFocus
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_GotFocus(lIndex As Integer)"
End Sub

Public Sub Event_Form_DblClick(lIndex As Integer)
On Local Error GoTo ErrHandler
ProcessSub lProject.pForms(lIndex).fSubs.sForm_DblClick
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_DblClick(lIndex As Integer)"
End Sub

Public Sub Event_Form_LostFocus(lIndex As Integer)
On Local Error GoTo ErrHandler
ProcessSub lProject.pForms(lIndex).fSubs.sForm_LostFocus
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_LostFocus(lIndex As Integer)"
End Sub

Public Sub Event_Form_Activate(lIndex As Integer)
On Local Error GoTo ErrHandler
MsgBox "Activate Sub is: " & vbCrLf & lProject.pForms(lIndex).fSubs.sForm_Activate
ProcessSub lProject.pForms(lIndex).fSubs.sForm_Activate
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_Activate(lIndex As Integer)"
End Sub

Public Sub Event_Form_Position(lIndex As Integer)
On Local Error GoTo ErrHandler
ProcessSub lProject.pForms(lIndex).fSubs.sForm_Position
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_Position(lIndex As Integer)"
End Sub

Public Sub Event_Form_Click(lIndex As Integer)
On Local Error GoTo ErrHandler
ProcessSub lProject.pForms(lIndex).fSubs.sForm_Click
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_Click(lIndex As Integer)"
End Sub

Public Sub Event_Form_Load(lIndex As Integer)
On Local Error GoTo ErrHandler
Dim msg() As String
ProcessSub lProject.pForms(lIndex).fSubs.sForm_Load
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_Load(lIndex As Integer)"
End Sub

Public Sub Event_Form_Objects(lIndex As Integer)
On Local Error GoTo ErrHandler
Dim msg() As String
'MsgBox "Form_Objects" & vbCrLf & vbCrLf & lProject.pForms(lIndex).fSubs.sForm_Objects
lCurrentForm = lIndex
ProcessSub lProject.pForms(lIndex).fSubs.sForm_Objects
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_Objects(lIndex As Integer)"
End Sub

Public Sub Event_Form_Resize(lIndex As Integer)
On Local Error GoTo ErrHandler
ProcessSub lProject.pForms(lIndex).fSubs.sForm_Resize
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub Event_Form_Resize(lIndex As Integer)"
End Sub

Public Sub SetFormLeft(lIndex As Integer, lLeft As Integer)
lProject.pForms(lIndex).fForm.Left = lLeft
End Sub

Public Sub SetFormTop(lIndex As Integer, lTop As Integer)
lProject.pForms(lIndex).fForm.Top = lTop
End Sub

Public Sub SetFormHeight(lIndex As Integer, lHeight As Integer)
lProject.pForms(lIndex).fForm.Height = lHeight
End Sub

Public Sub SetFormWidth(lIndex As Integer, lWidth As Integer)
lProject.pForms(lIndex).fForm.Width = lWidth
End Sub

Public Function ReturnFormIndex(lName As String) As Integer
On Local Error GoTo ErrHandler
Dim i As Integer
'MsgBox lName & " - " & lProject.pFormCount
If Len(lName) <> 0 Then
    For i = 0 To lProject.pFormCount
'        MsgBox "Comparing " & LCase(Trim(lProject.pForms(i).fName)) & " -- " & LCase(Trim(lName))
        If LCase(Trim(lProject.pForms(i).fName)) = LCase(Trim(lName)) Then
'            MsgBox "FOUND IT " & i
            ReturnFormIndex = i
            Exit For
        End If
    Next i
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReturnFormIndex(lName As String) As Integer"
End Function

Public Function ProcessSub(lSub As String) As Boolean
MsgBox "Processing Sub: " & vbCrLf & lSub
On Local Error GoTo ErrHandler
Dim lSubName As String
Dim i As Integer, l As Integer, msg() As String, msg2 As String, msg3 As String, msg4 As String, b As Boolean, f As Integer, splt() As String, splt2() As String, t As Integer
Dim lStringCount As Integer, lStrings(128) As String, lStringNames(128) As String
Dim lIntegerCount As Integer, lIntegers(128) As Integer, lIntegerNames(128) As String
msg = Split(lSub, vbCrLf)
For i = 0 To UBound(msg)
    If Len(msg(i)) <> 0 Then
        'If Left(Trim(msg(i)), 6) = "object" Then
'            MsgBox "OBJECT"
'            'Set splt2 = Nothing
'            'msg2 = Right(msg(i), Len(msg(i)) - 7)
'            splt2 = Split(Right(msg(i), Len(msg(i)) - 7), ", ")
'            MsgBox msg(i)
'            For l = 0 To UBound(splt2)
'                MsgBox splt2(l)
'            Next l
'            'msg2 = ""
''        End If
        
        msg(i) = Trim(msg(i))
        If Left(LCase(Trim(msg(i))), 8) = "position" Then
            'MsgBox "Found Position Statement"
            msg3 = Right(msg(i), Len(msg(i)) - 9)
            msg3 = Left(msg3, Len(msg3) - 1)
            splt2 = Split(msg3, ",")
            For l = 0 To UBound(splt2)
                Select Case l
                Case 0
'                    MsgBox "MAI_ReturnFormIndex: " & splt2(l)
                    t = ReturnFormIndex(splt2(l))
'                    MsgBox "FormIndex: " & t
                Case 1
                    If t <> 0 Then SetFormLeft t, Int(Trim(splt2(l)))
                Case 2
                    If t <> 0 Then SetFormTop t, Int(Trim(splt2(l)))
                Case 3
                    If t <> 0 Then SetFormWidth t, Int(Trim(splt2(l)))
                Case 4
                    If t <> 0 Then SetFormHeight t, Int(Trim(splt2(l)))
                End Select
            Next l
            'splt2(l) = Replace(splt2(l), ")", "")
        End If
        If LCase(Trim(msg(i))) = "end" Then End
        If Left(LCase(msg(i)), 7) = "object " Then
            If lCurrentForm <> 0 Then
                'MsgBox msg(i)
                msg3 = Right(msg(i), Len(msg(i)) - 6)
                msg4 = Parse(msg3, Left(msg3, 2), "(")
                msg3 = Right(msg3, Len(msg3) - Len(msg4))
                splt2 = Split(Right(msg3, Len(msg3) - 7), ", ")
                Select Case LCase(Trim(msg4))
                Case "button"
                    For l = 0 To UBound(splt2)
                        With lProject.pForms(lCurrentForm).fForm
                            splt2(l) = Replace(splt2(l), ")", "")
                            Select Case l
                            Case 0
                                t = .cmdButton.Count + 1
                                Load .cmdButton(t)
                                .cmdButton(t).Tag = splt2(1)
                                .cmdButton(t).Visible = True
                            Case 1
                                .cmdButton(t).Caption = splt2(l)
                            Case 2
                                .cmdButton(t).Left = Int(splt2(l))
                            Case 3
                                .cmdButton(t).Top = Int(splt2(l))
                            Case 4
                                .cmdButton(t).Width = Int(splt2(l))
                            Case 5
                                .cmdButton(t).Height = Int(splt2(l))
                            End Select
                        End With
                        'MsgBox splt2(l)
                    Next l
                End Select
            End If
        End If
        If Left(LCase(msg(i)), 7) = "unload " Then
            msg2 = Right(msg(i), Len(msg(i)) - 7)
            l = ReturnFormIndex(msg2)
            If l <> 0 Then Unload lProject.pForms(l).fForm
        End If
        If InStr(LCase(msg(i)), ".hide") Then
            msg2 = Left(msg(i), Len(msg(i)) - 5)
            For l = 1 To lProject.pFormCount
                If Trim(LCase(lProject.pForms(l).fName)) = Trim(LCase(msg2)) Then
                    lProject.pForms(l).fForm.Visible = False
                    Exit For
                End If
            Next l
        End If
        If InStr(LCase(msg(i)), ".show") Then
            msg2 = Left(msg(i), Len(msg(i)) - 5)
            For l = 1 To lProject.pFormCount
                If Trim(LCase(lProject.pForms(l).fName)) = Trim(LCase(msg2)) Then
                    lProject.pForms(l).fForm.Visible = False
                    lProject.pForms(l).fForm.Show
                    lProject.pForms(l).fForm.WindowState = vbNormal
                    lProject.pForms(l).fName = Trim(Str(l))
                    'lProject.pForms(l).fForm.Tag = Trim(Str(l))
                    Exit For
                End If
            Next l
        End If
        If InStr(msg(i), " = ") Then
            splt = Split(msg(i), "=")
            Select Case UBound(splt)
            Case 1
                For l = 0 To lStringCount
                    If Trim(LCase(splt(0))) = Trim(LCase(lStringNames(l))) Then
                        lStrings(l) = Replace(Trim(splt(1)), Chr(34), "")
                    End If
                Next l
                For l = 0 To lIntegerCount
                    If Trim(LCase(splt(0))) = Trim(LCase(lIntegerNames(l))) Then
                        If InStr(splt(1), "-") Then
                            splt2 = Split(splt(1), " - ")
                            t = Trim(Int(splt2(0))) - Trim(Int(splt2(1)))
                            lIntegers(l) = t
                            Exit For
                        End If
                        If InStr(splt(1), "+") Then
                            splt2 = Split(splt(1), " + ")
                            For f = 0 To UBound(splt2)
                                t = t + Int(Trim(splt2(f)))
                            Next f
                            lIntegers(l) = t
                            Exit For
                        End If
                    End If
                Next l
            Case Else
                RaiseEvent RuntimeError(rTooManyParameters)
            End Select
        End If
        If Left(Trim(LCase(msg(i))), 6) = "msgbox" Then
            msg3 = Right(msg(i), Len(msg(i)) - 7)
            'MsgBox "REACHED MSGBOX " & vbCrLf & msg3
            If Left(Trim(msg3), 1) = Chr(34) And Right(Trim(msg3), 1) = Chr(34) Then
                MsgBox Replace(msg3, Chr(34), "")
                GoTo EndMsgBox
            End If
            For l = 0 To lIntegerCount
                If LCase(Trim(lIntegerNames(l))) = LCase(Trim(msg3)) Then
                    MsgBox lIntegers(l), , lProject.pName
                    Exit For
                End If
            Next l
            For l = 0 To lStringCount
                If LCase(Trim(lStringNames(l))) = LCase(Trim(msg3)) Then
                    MsgBox lStrings(l), , lProject.pName
                    Exit For
                End If
            Next l
        End If
EndMsgBox:
        If Left(Trim(LCase(msg(i))), 3) = "dim" Then
            If InStr(msg(i), ", ") Then
                msg2 = ""
                splt2 = Split(msg(i), ", ")
                For l = 0 To UBound(splt2)
                    If Len(splt2(l)) <> 0 Then
                        splt2(l) = Replace(splt2(l), "Dim ", "")
                        splt2(l) = Replace(splt2(l), "dim ", "")
                        For t = 0 To Len(splt2(l))
                            If Asc(Left(splt2(l), 1)) <> 32 Then
                                If Len(msg2) <> 0 Then
                                    msg2 = msg2 & Left(splt2(l), 1)
                                Else
                                    msg2 = Left(splt2(l), 1)
                                End If
                                splt2(l) = Right(splt2(l), Len(splt2(l)) - 1)
                            Else
                                splt2(l) = Right(splt2(l), Len(splt2(l)) - 1)
                                Select Case LCase(Trim(splt2(l)))
                                Case "as integer"
                                    lIntegerCount = lIntegerCount + 1
                                    lIntegers(lIntegerCount) = 0
                                    lIntegerNames(lIntegerCount) = msg2
                                    msg2 = ""
                                Case "as string"
                                    lStringCount = lStringCount + 1
                                    lStrings(lStringCount) = ""
                                    lStringNames(lStringCount) = msg2
                                    msg2 = ""
                                    MsgBox lStringNames(lStringCount)
                                End Select
                                Exit For
                            End If
                        Next t
                        
                    End If
                Next l
            Else
                msg2 = ""
                msg(i) = Right(msg(i), Len(msg(i)) - 4)
                For l = 0 To Len(msg(i))
                    If Asc(Left(msg(i), 1)) <> 32 Then
                        If Len(msg2) <> 0 Then
                            msg2 = msg2 & Left(msg(i), 1)
                        Else
                            msg2 = Left(msg(i), 1)
                        End If
                        msg(i) = Right(msg(i), Len(msg(i)) - 1)
                    Else
                        msg(i) = Right(msg(i), Len(msg(i)) - 1)
                        Select Case LCase(Trim(msg(i)))
                        Case "as integer"
                            lIntegerCount = lIntegerCount + 1
                            lIntegers(lIntegerCount) = 0
                            lIntegerNames(lIntegerCount) = msg2
                        Case "as string"
                            lStringCount = lStringCount + 1
                            lStrings(lStringCount) = ""
                            lStringNames(lStringCount) = msg2
                        Case "as long"
                        Case "as double"
                        Case Else
                        End Select
                        Exit For
                    End If
                Next l
            End If
        End If
        If Left(Trim(LCase(msg(i))), 3) = "sub" Then
            lSubName = Right(msg(i), Len(msg(i)) - 3)
            lSubName = Left(lSubName, Len(lSubName) - Len(Parse(lSubName, "(", ")")) - 2)
        End If
    End If
Next i
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ProcessSub(lSub As String) As Boolean"
End Function
