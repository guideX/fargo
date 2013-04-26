Attribute VB_Name = "mdlProfileStrings"
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function ReadINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, Optional lDefault As String)
On Local Error GoTo ErrHandler
Dim msg As String, RetVal As String, Worked As Integer
RetVal = String$(255, 0)
Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), lFile)
If Worked = 0 Then
    ReadINI = lDefault
Else
    ReadINI = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
End If
Exit Function
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Function ReadINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, Optional lDefault As String)"
End Function

Public Sub WriteINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
On Local Error GoTo ErrHandler
WritePrivateProfileString Section, Key, Value, lFile
Exit Sub
ErrHandler:
    ProcessRuntimeError Err.Number, Err.Description, "Public Sub WriteINI(ByVal lFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)"
End Sub

