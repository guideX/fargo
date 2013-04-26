Attribute VB_Name = "mdlGeneral"
Option Explicit

Public Function SaveFile(lFileName As String, lText As String) As Boolean
On Local Error GoTo ErrHandler
If Len(lFileName) <> 0 And Len(lText) <> 0 Then
    Open lFileName For Output As #1
    Print #1, lText
    Close #1
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function SaveFile(lFileName As String, lText As String) As Boolean"
End Function

Public Function GetFileTitle(lFileName As String) As String
On Local Error GoTo ErrHandler
Dim msg() As String
If Len(lFileName) <> 0 Then
    msg = Split(lFileName, "\", -1, vbTextCompare)
    GetFileTitle = msg(UBound(msg))
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function GetFileTitle(lFileName As String) As String"
End Function

Public Function Parse(lWhole As String, lStart As String, lEnd As String)
On Local Error GoTo ErrHandler
Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
If Len(Trim(lStart)) <> 0 And Len(Trim(lEnd)) <> 0 Then
    len1 = InStr(lWhole, lStart)
    len2 = InStr(lWhole, lEnd)
    Str1 = Right(lWhole, Len(lWhole) - len1)
    Str2 = Right(lWhole, Len(lWhole) - len2)
    Parse = Left(Str1, Len(Str1) - Len(Str2) - 1)
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function Parse(lWhole As String, lStart As String, lEnd As String)"
End Function

Public Function DoesFileExist(lFileName As String) As Boolean
On Local Error GoTo ErrHandler
Dim msg As String
msg = Dir(lFileName)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function DoesFileExist(lFileName As String) As Boolean"
End Function

Public Function ReadFile(lFile As String) As String
On Local Error GoTo ErrHandler
Dim n As Integer, msg As String
n = FreeFile
If Len(lFile) <> 0 Then
    Open lFile For Input As #n
        msg = StrConv(InputB(LOF(n), n), vbUnicode)
        If Len(msg) <> 0 Then
            ReadFile = Left(msg, Len(msg) - 2)
        End If
    Close #n
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReadFile(lFile As String) As String"
End Function

Public Function ReturnFilePath(lFile As String) As String
On Local Error GoTo ErrHandler
ReturnFilePath = Left(lFile, Len(lFile) - Len(GetFileTitle(lFile)))
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReadFile(lFile As String) As String"
End Function
