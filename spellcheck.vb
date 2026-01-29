
Option Explicit

' Entry point: spellcheck Column A (rows 2..last used row) on the active worksheet.
Public Sub Spellcheck_ColumnA_CurrentSheet()
	Dim ws As Worksheet
	Set ws = ActiveSheet
	If ws Is Nothing Then Exit Sub

	Dim lastRow As Long
	lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
	If lastRow < 2 Then Exit Sub

	Dim rng As Range
	Set rng = ws.Range("A2:A" & lastRow)

	' Pre-scan: detect if any misspellings exist so we can decide whether to show the follow-up prompts.
	Dim hadSpellingErrors As Boolean
	hadSpellingErrors = RangeHasSpellingErrors(rng)

	' Run Excel's built-in spellcheck dialog over the range.
	On Error GoTo CleanFail
	rng.CheckSpelling

	If hadSpellingErrors Then
		PromptUserConfirmationLoop
	End If

	Exit Sub

CleanFail:
	' If spellcheck is canceled or fails for any reason, just exit quietly.
	Exit Sub
End Sub

Private Sub PromptUserConfirmationLoop()
	Dim answer As VbMsgBoxResult
	Dim sure As VbMsgBoxResult

	Do
		answer = MsgBox("Have you checked the spelling?", vbQuestion Or vbYesNo, "Spellcheck")
		If answer = vbNo Then
			MsgBox "You should", vbExclamation, "Spellcheck"
		Else
			' They said Yes; now confirm.
			sure = MsgBox("Are you sure?", vbQuestion Or vbYesNo, "Spellcheck")
			If sure = vbYes Then
				MsgBox "Okay then I believe you", vbInformation, "Spellcheck"
				Exit Do
			End If
		End If
	Loop
End Sub

Private Function RangeHasSpellingErrors(ByVal rng As Range) As Boolean
	Dim cell As Range
	For Each cell In rng.Cells
		If Not IsError(cell.Value2) Then
			Dim textValue As String
			textValue = Trim$(CStr(cell.Value2))
			If Len(textValue) > 0 Then
				If TextHasSpellingErrors(textValue) Then
					RangeHasSpellingErrors = True
					Exit Function
				End If
			End If
		End If
	Next cell

	RangeHasSpellingErrors = False
End Function

Private Function TextHasSpellingErrors(ByVal textValue As String) As Boolean
	Dim cleaned As String
	cleaned = NormalizeForWordSplit(textValue)

	Dim parts As Variant
	parts = Split(cleaned, " ")

	Dim i As Long
	For i = LBound(parts) To UBound(parts)
		Dim w As String
		w = Trim$(CStr(parts(i)))
		If Len(w) > 1 And ContainsLetter(w) Then
			If Not Application.CheckSpelling(w) Then
				TextHasSpellingErrors = True
				Exit Function
			End If
		End If
	Next i

	TextHasSpellingErrors = False
End Function

Private Function NormalizeForWordSplit(ByVal s As String) As String
	' Replace common punctuation with spaces so Split() finds words.
	Dim t As String
	t = s

	' Convert line breaks/tabs to spaces
	t = Replace(t, vbCr, " ")
	t = Replace(t, vbLf, " ")
	t = Replace(t, vbTab, " ")

	' Punctuation to spaces
	Dim punct As Variant
	punct = Array(".", ",", ";", ":", "!", "?", "(", ")", "[", "]", "{", "}", """""", "'", "“", "”", "‘", "’", "-", "–", "—", "/", "\", "|", "_", "=", "+", "*", "&", "^", "%", "#", "@", "~", "`")

	Dim i As Long
	For i = LBound(punct) To UBound(punct)
		t = Replace(t, CStr(punct(i)), " ")
	Next i

	' Collapse multiple spaces
	Do While InStr(t, "  ") > 0
		t = Replace(t, "  ", " ")
	Loop

	NormalizeForWordSplit = Trim$(t)
End Function

Private Function ContainsLetter(ByVal s As String) As Boolean
	Dim i As Long
	For i = 1 To Len(s)
		Dim ch As Integer
		ch = AscW(Mid$(s, i, 1))
		If (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Then
			ContainsLetter = True
			Exit Function
		End If
	Next i
	ContainsLetter = False
End Function

