VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lErrorCount As Integer

Public Sub ProcessError(lNumber As Long, lDescription As String, lSub As String)
Dim l As Long
lErrorCount = lErrorCount + 1
l = CLng(ReadINI(App.Path & "\DATA\CONFIG\ERRORLOG.INI", lSub, "Quanity", 0)) + 1
WriteINI App.Path & "\DATA\CONFIG\ERRORLOG.INI", lSub, "Quanity", l
WriteINI App.Path & "\DATA\CONFIG\ERRORLOG.INI", lSub, "Number", Trim(Str(lNumber))
WriteINI App.Path & "\DATA\CONFIG\ERRORLOG.INI", lSub, "Description", lDescription
MsgBox "Error: " & vbCrLf & "Sub: " & lSub & vbCrLf & "Description: " & lDescription
End Sub
