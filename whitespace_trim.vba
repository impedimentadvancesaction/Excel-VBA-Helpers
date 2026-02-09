Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitPoint

    'Limit to columns A:G
    If Intersect(Target, Me.Range("A:G")) Is Nothing Then Exit Sub

    'Ignore multi-cell pastes
    If Target.CountLarge > 1 Then Exit Sub

    If Not IsError(Target.Value) Then
        Dim v As String
        v = CStr(Target.Value)

        'Remove tabs and line breaks
        v = Replace(v, vbTab, "")
        v = Replace(v, vbCr, "")
        v = Replace(v, vbLf, "")

        'Remove non-breaking spaces
        v = Replace(v, Chr(160), " ")

        'Trim leading/trailing spaces
        v = Trim(v)

        'Prevent recursion
        Application.EnableEvents = False
        Target.Value = v
        Application.EnableEvents = True
    End If

ExitPoint:
    Application.EnableEvents = True
End Sub
