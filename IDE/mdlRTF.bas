Attribute VB_Name = "mdlRTF"
'Option Explicit
Private lCurrentCharecter As String
Private lThisLine As String
Private lStart As String
Private lTend As String
Private lHoldTend As String
Private lHoldStart As String
Private lTopLine As String
Private lFoundPos As String
Private lCommentCharecter As String
Private lLongVar As String
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Sub SetRTFColors(lRTF As RichTextBox, lhWnd As Long)
lCommentCharecter = "'"
lLongVar = "$"
'If KeyCode = 13 Then Exit Sub
LockWindowUpdate lhWnd
clearwordcolors lRTF
ColorizeWord lRTF, lLongVar, &H80&
ColorizeWord lRTF, lCommentCharecter, &H8000&
ColorizeWord lRTF, "or", &H800000
ColorizeWord lRTF, "and", &H800000
ColorizeWord lRTF, "random", &H800000
ColorizeWord lRTF, "append", &H800000
ColorizeWord lRTF, "binary", &H800000
ColorizeWord lRTF, "exit", &H800000
ColorizeWord lRTF, "then", &H800000
ColorizeWord lRTF, "goto", &H800000
ColorizeWord lRTF, "case", &H800000
ColorizeWord lRTF, "select", &H800000
ColorizeWord lRTF, "end", &H800000
ColorizeWord lRTF, "Select Case", &H800000
ColorizeWord lRTF, "End select", &H800000
ColorizeWord lRTF, "for", &H800000
ColorizeWord lRTF, "each", &H800000
ColorizeWord lRTF, "loop", &H800000
ColorizeWord lRTF, "While", &H800000
ColorizeWord lRTF, "Until", &H800000
ColorizeWord lRTF, "for each", &H800000
ColorizeWord lRTF, "Next", &H800000
ColorizeWord lRTF, "True", &H800000
ColorizeWord lRTF, "False", &H800000
ColorizeWord lRTF, "sub", &H800000
ColorizeWord lRTF, "function", &H800000
ColorizeWord lRTF, "Integer", &H800000
ColorizeWord lRTF, "As", &H800000
ColorizeWord lRTF, "Private", &H800000
ColorizeWord lRTF, "Dim", &H800000
ColorizeWord lRTF, "MsgBox", &H800000
ColorizeWord lRTF, "else", &H800000
ColorizeWord lRTF, "else if", &H800000
ColorizeWord lRTF, "Public", &H800000
ColorizeWord lRTF, "Close", &H800000
ColorizeWord lRTF, "Open", &H800000
ColorizeWord lRTF, "End If", &H800000
ColorizeWord lRTF, "If", &H800000
ColorizeWord lRTF, "(", &H800000
ColorizeWord lRTF, ")", &H800000
LockWindowUpdate 0&
lRTF.Enabled = True
If lRTF.Visible = True Then
lRTF.SetFocus
End If
End Sub

Private Function ColorizeWord(Rich1 As RichTextBox, Word As String, Color As OLE_COLOR)
Dim lStartLine As String, lNowLine As String, commentposx
Do Until Rich1.GetLineFromChar(lStart) <> lThisLine
    lStart = lStart - 1
    If lStart < 0 Then
        lStart = 0
        Exit Do
    End If
Loop
lStartLine = Rich1.GetLineFromChar(Rich1.SelStart)
If Rich1.SelLength > 0 Then Exit Function
Rich1.Enabled = False
lStart = lStart
If lStart = 0 Then
    lStart = 1
End If
lStart = lStart - Len(Word)
Do
lNowLine = Rich1.GetLineFromChar(Rich1.SelStart)
If lNowLine <> lStartLine Then GoTo endx
lHoldStart = lStart + Len(Word)
commentposx = InStr(lHoldStart, Rich1.Text, lCommentCharecter, vbTextCompare)
If lHoldStart < 1 Then
    lHoldStart = 1
End If
lStart = lStart + Len(Word)
lFoundPos = InStr(lStart, Rich1.Text, Word, vbTextCompare)
If lFoundPos > lTend Then GoTo endx '''''''''''''''''''''
If lFoundPos < 1 Then GoTo endx
If lFoundPos < 2 Then
sletter = ""
Else
sletter = Mid(Rich1.Text, lFoundPos - 1, 1)
End If
eletter = Mid(Rich1.Text, lFoundPos + Len(Word), 1)
If lFoundPos > 0 Then
If lFoundPos = 1 Then
lStart = lStart - 1
End If
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
If Word = lCommentCharecter Then
 lTend = Rich1.SelStart
       Do Until Rich1.GetLineFromChar(lTend) <> lThisLine
            lTend = lTend + 1
            If lTend > Len(Rich1.Text) Then
                lTend = Len(Rich1.Text) + 1
                Exit Do
            End If
        Loop
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = lTend - (lFoundPos - 1)
Rich1.SelColor = Color
Rich1.SelLength = 0
Rich1.SelStart = lCurrentCharecter
Rich1.SelColor = &H0&
Exit Function
Exit Do
End If
''''''''''''''''''''''''''''''
If Word = lLongVar Then
 lTend = Rich1.SelStart
       Do Until Rich1.GetLineFromChar(lTend) <> lThisLine
            lTend = lTend + 1
            If lTend > Len(Rich1.Text) Then
                lTend = Len(Rich1.Text) + 1
                Exit Do
            End If
        Loop

pos = lStart
Do
lFoundPos = InStr(pos, Rich1.Text, lLongVar, vbTextCompare)

For i = lFoundPos To lTend
If lFoundPos < 1 Then Exit For
If i = lTend Then Exit For
Rich1.SelStart = i - 1
Rich1.SelLength = 1
If Rich1.SelText = "" Then Exit For
Select Case Asc(Rich1.SelText)
Case 48 To 57
Rich1.SelColor = Color
Case 36
Rich1.SelColor = Color
Case 97 To 122
Rich1.SelColor = Color
Case 65 To 90
Rich1.SelColor = Color
Case 145
Rich1.SelColor = Color
Case 146
Rich1.SelColor = Color
Case 143
Rich1.SelColor = Color
Case 143
Rich1.SelColor = Color
Case Else
Exit For
End Select

Next

pos = lFoundPos + 2
Loop While lFoundPos > 0

GoTo endx
End If


If lStart = 0 Then
lStart = 1
End If
commentposx = InStr(lStart, Rich1.Text, lCommentCharecter, vbTextCompare)
If commentposx > 0 Then
If Rich1.SelStart > commentposx Then GoTo endx
End If

If Len(Word) = 1 Then
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If



If eletter = "" And sletter = "" Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If
If eletter = "" And sletter = " " Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If
If eletter = " " And sletter = "" Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If
If eletter = " " And sletter = " " Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If
If eletter = "" And sletter = Chr(10) Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If

If eletter = " " And sletter = Chr(10) Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If

If eletter = Chr(10) And sletter = "" Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If


If eletter = Chr(10) And sletter = " " Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If


If eletter = Chr(10) And sletter = Chr(10) Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If

If eletter = Chr(13) And sletter = Chr(10) Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If

If eletter = Chr(13) And sletter = "" Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If

If eletter = Chr(13) And sletter = " " Then
theword = Rich1.SelText
originaltext = Rich1.SelText
theword = LCase(theword)
firstchar = Mid(theword, 1, 1)
rest = Mid(theword, 2, Len(theword))
firstchar = UCase(firstchar)
Rich1.SelText = firstchar & rest
Rich1.SelStart = lFoundPos - 1
Rich1.SelLength = Len(Word)
Rich1.SelColor = Color
End If

Rich1.SelLength = 0
End If
If lFoundPos = 1 Then
lStart = lStart + 1
End If

Loop While lFoundPos > 0
endx:



Rich1.SelStart = lCurrentCharecter
Rich1.SelColor = &H0&
lFoundPos = 0
eletter = ""
sletter = ""
Rich1.Enabled = True

End Function

Public Function clearwordcolors(Rich1 As RichTextBox)
If Rich1.SelLength > 0 Then Exit Function
Rich1.Enabled = False
lCurrentCharecter = Rich1.SelStart
lThisLine = Rich1.GetLineFromChar(Rich1.SelStart)
lStart = Rich1.SelStart
lTend = Rich1.SelStart
With Rich1
      Do Until .GetLineFromChar(lStart) <> lThisLine
            lStart = lStart - 1
            If lStart < 0 Then
                lStart = 0
                Exit Do
            End If
        
        Loop



       Do Until .GetLineFromChar(lTend) <> lThisLine
            lTend = lTend + 1
            If lTend > Len(.Text) Then
                lTend = Len(.Text) + 1
                Exit Do
            End If

Loop
End With
If lStart = 1 Then
lTend = lTend - 2
End If
If lStart > 1 Then
lStart = lStart + 1
lTend = lTend - 1
End If
lHoldStart = lStart
lHoldTend = lTend
Rich1.SelStart = lStart
Rich1.SelLength = lTend - lStart
Rich1.SelColor = &H0&
Rich1.SelLength = 0
Rich1.SelStart = lCurrentCharecter
lHoldTend = lTend
Rich1.Enabled = True
End Function
