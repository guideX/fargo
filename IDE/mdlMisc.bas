Attribute VB_Name = "mdlMisc"
Option Explicit
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Function SaveDialog(lForm As Form, lFilter As String, lCaption As String, lStartingPath As String) As String
mdiMain.CommonDialog1.Filter = lFilter
mdiMain.CommonDialog1.DialogTitle = lCaption
mdiMain.CommonDialog1.InitDir = lStartingPath
mdiMain.CommonDialog1.ShowSave
SaveDialog = mdiMain.CommonDialog1.FileName
End Function

Public Sub FormDrag(lFormName As Form)
ReleaseCapture
Call SendMessage(lFormName.hWnd, &HA1, 2, 0&)
End Sub

Public Function Parse(lWhole As String, lStart As String, lEnd As String)
On Local Error GoTo ErrHandler
Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
len1 = InStr(lWhole, lStart)
len2 = InStr(lWhole, lEnd)
Str1 = Right(lWhole, Len(lWhole) - len1)
Str2 = Right(lWhole, Len(lWhole) - len2)
Parse = Left(Str1, Len(Str1) - Len(Str2) - 1)
ErrHandler:
End Function

Public Sub CutRegion(ctlHwnd As Long, Ctl As ComboBox, bCut As Boolean)
Dim hRgn As Long
If bCut = True Then
    hRgn = CreateRectRgn(1, 1, ((Ctl.Width / Screen.TwipsPerPixelX) - 3), ((Ctl.Height / Screen.TwipsPerPixelY) - 3))
Else
    hRgn = CreateRectRgn(0, 0, (Ctl.Width / Screen.TwipsPerPixelX), (Ctl.Height / Screen.TwipsPerPixelY))
End If
SetWindowRgn ctlHwnd, hRgn, True
End Sub

Public Function DoesFileExist(lFileName As String) As Boolean
Dim msg As String
msg = Dir(lFileName)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Function GetFileTitle(lFileName As String) As String
Dim msg() As String
If Len(lFileName) <> 0 Then
    msg = Split(lFileName, "\", -1, vbTextCompare)
    GetFileTitle = msg(UBound(msg))
End If
End Function

Public Function SaveFile(lFileName As String, lText As String) As Boolean
If Len(lFileName) <> 0 And Len(lText) <> 0 Then
    Open lFileName For Output As #1
    Print #1, lText
    Close #1
End If
End Function

Public Function ReadFile(lFile As String) As String
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
End Function

