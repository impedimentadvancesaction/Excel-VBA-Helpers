Option Explicit

'================================================================================
' CHECK NOTIFICATION WORDING
' Highlights banned terms (from Data sheet) and repeated words in notification
' cells. Banned single words and phrases = red; repeated (non-banned) words = orange.
'================================================================================

Public Sub CheckNotificationWording()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Load banned terms from Data sheet column A (row 2 to last used).
    Dim banned As Object
    Set banned = LoadBannedTerms(ThisWorkbook.Worksheets("Data"), "A", 2)

    ' Run token-level highlighting first (banned words + repeats), then phrase
    ' highlighting last so phrase red is not wiped by the cell reset to black.
    AnalyseAndHighlightCell ws.Range("A1"), banned
    AnalyseAndHighlightCell ws.Range("A4"), banned

    HighlightBannedPhrases ws.Range("A1"), banned
    HighlightBannedPhrases ws.Range("A4"), banned

    ForceSpellCheck ws

    MsgBox "Wording check complete.", vbInformation

End Sub

'============================================================
' Run spell check on all of column A (used range) after formatting.
'============================================================
Private Sub ForceSpellCheck(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then Exit Sub
    ws.Range(ws.Cells(1, "A"), ws.Cells(lastRow, "A")).CheckSpelling
End Sub

'============================================================
' Load banned terms (case-sensitive) from Data!A2:A(last)
'============================================================
Private Function LoadBannedTerms(ByVal wsData As Worksheet, ByVal colLetter As String, ByVal startRow As Long) As Object
    ' Returns Dictionary: key = banned term (cleaned), item = True if phrase else False.

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 0 ' Binary (case-sensitive)

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, colLetter).End(xlUp).Row
    If lastRow < startRow Then
        Set LoadBannedTerms = dict
        Exit Function
    End If

    Dim c As Range
    Dim key As String

    For Each c In wsData.Range(wsData.Cells(startRow, colLetter), wsData.Cells(lastRow, colLetter))
        key = CleanToken(c.Value2)
        If Len(key) > 0 Then
            dict.Add key, InStr(key, " ") > 0   ' True = phrase
        End If
    Next c

    Set LoadBannedTerms = dict

End Function

'============================================================
' Tokenise + highlight: red banned, orange repeated (non-banned)
'============================================================
Private Sub AnalyseAndHighlightCell(ByVal cell As Range, ByVal banned As Object)
    ' Single-word banned terms = red; repeated (non-banned) words = orange; rest = black.
    If cell.MergeCells Then Exit Sub

    Dim txt As String
    txt = CStr(cell.Value2)
    If Len(txt) = 0 Then Exit Sub

    ' Force rich-text mode and reset to black
    cell.Value = txt
    cell.Characters.Font.Color = vbBlack

    ' Tokenise text into words with their exact start positions and lengths
    Dim starts As Collection, lens As Collection, tokens As Collection
    Set starts = New Collection
    Set lens = New Collection
    Set tokens = New Collection

    Tokenise txt, starts, lens, tokens

    ' Count how many times each token appears (case-sensitive)
    Dim counts As Object
    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = 0 ' Binary compare (case-sensitive)

    Dim i As Long, tok As String
    For i = 1 To tokens.Count
        tok = tokens(i)
        If counts.Exists(tok) Then
            counts(tok) = counts(tok) + 1
        Else
            counts.Add tok, 1
        End If
    Next i

    ' Apply formatting per token (never color anything that isn't a token)
    Dim st As Long, ln As Long
    For i = 1 To tokens.Count
        tok = tokens(i)
        st = CLng(starts(i))
        ln = CLng(lens(i))

        If banned.Exists(tok) Then
            cell.Characters(Start:=st, length:=ln).Font.Color = vbRed
        ElseIf counts(tok) > 1 Then
            cell.Characters(Start:=st, length:=ln).Font.Color = RGB(255, 165, 0) ' orange
        End If
    Next i
End Sub

'============================================================
' Tokeniser: defines a "word" as A-Z a-z 0-9 or hyphen "-"
' Keeps exact start/length so highlighting is always correct.
'============================================================
Private Sub Tokenise(ByVal txt As String, ByVal starts As Collection, ByVal lens As Collection, ByVal tokens As Collection)
    ' Fills starts (1-based start index), lens (length), tokens (cleaned word) per token.
    Dim i As Long, ch As String
    Dim inTok As Boolean
    Dim tokStart As Long
    Dim tok As String

    inTok = False
    tok = vbNullString

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)

        If IsTokenChar(ch) Then
            If Not inTok Then
                inTok = True
                tokStart = i
                tok = ch
            Else
                tok = tok & ch
            End If
        Else
            If inTok Then
                starts.Add tokStart
                lens.Add Len(tok)
                tokens.Add CleanToken(tok)
                inTok = False
                tok = vbNullString
            End If
        End If
    Next i

    ' Don't forget the last token if text doesn't end with a non-token char
    If inTok Then
        starts.Add tokStart
        lens.Add Len(tok)
        tokens.Add CleanToken(tok)
    End If
End Sub

' Returns True if ch is a letter, digit, or hyphen (part of a token).
Private Function IsTokenChar(ByVal ch As String) As Boolean
    IsTokenChar = (ch Like "[A-Za-z0-9]") Or (ch = "-")
End Function

'============================================================
' Normalise Unicode dashes/spaces but keep case. Trims result.
' ChrW(8211)=en-dash, 8212=em-dash, 8209=non-breaking hyphen, 173=soft hyphen;
' Chr(160)=non-breaking space.
'============================================================
Private Function CleanToken(ByVal v As Variant) As String
    Dim t As String
    t = CStr(v)

    t = Replace(t, ChrW(8211), "-")
    t = Replace(t, ChrW(8212), "-")
    t = Replace(t, ChrW(8209), "-")
    t = Replace(t, ChrW(173), "-")
    t = Replace(t, Chr(160), " ")

    CleanToken = Trim$(t)
End Function

' Same replacements as CleanToken but no Trim; preserves length for position mapping.
Private Function NormaliseForMatch(ByVal s As String) As String
    Dim t As String
    t = CStr(s)
    t = Replace(t, ChrW(8211), "-")
    t = Replace(t, ChrW(8212), "-")
    t = Replace(t, ChrW(8209), "-")
    t = Replace(t, ChrW(173), "-")
    t = Replace(t, Chr(160), " ")
    NormaliseForMatch = t
End Function

'============================================================
' Highlight every occurrence of banned phrases (multi-word terms) in red.
' Uses same-length normalised text so Character start/length match the cell.
'============================================================
Private Sub HighlightBannedPhrases(ByVal cell As Range, ByVal banned As Object)

    Dim cellText As String
    cellText = CStr(cell.Value)
    If Len(cellText) = 0 Then Exit Sub

    ' Search in same-length normalised copy so positions match the cell.
    Dim searchText As String
    searchText = NormaliseForMatch(cellText)

    Dim key As Variant
    Dim pos As Long
    Dim startPos As Long

    For Each key In banned.Keys
        If banned(key) = True Then   ' phrase
            startPos = 1
            Do
                pos = InStr(startPos, searchText, key, vbBinaryCompare)
                If pos = 0 Then Exit Do

                cell.Characters(Start:=pos, length:=Len(key)).Font.Color = vbRed
                startPos = pos + Len(key)
            Loop
        End If
    Next key

End Sub
