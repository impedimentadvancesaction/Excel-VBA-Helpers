Private Sub Worksheet_Change(ByVal Target As Range)

    ' This macro runs automatically whenever a cell on this worksheet is changed.
    ' The code below reacts to changes in two specific cells:
    '   - A2 (which controls whether A8 is restricted or free text)
    '   - A8 (to tidy up formatting after the user makes a selection)

    ' =========================================================
    ' PART 1: What happens when cell A2 is changed
    ' =========================================================
    If Not Intersect(Target, Me.Range("A2")) Is Nothing Then

        ' This list contains the values in A2 that should trigger
        ' a dropdown list in cell A8.
        ' IMPORTANT: These values are cleaned before comparison so that
        ' hidden characters (fake hyphens, non-breaking spaces, etc.)
        ' do not prevent a match.
        Dim specialList As Variant
        specialList = Array("IG-ND-21", "Value2", "Value3")  ' <-- Replace with your real trigger values

        Dim trigger As Boolean
        Dim v As Variant

        ' Check whether the value entered into A2 matches
        ' any value in the trigger list (ignoring case and hidden characters).
        For Each v In specialList
            If StrComp(CleanText(Target.Value), CleanText(v), vbTextCompare) = 0 Then
                trigger = True
                Exit For
            End If
        Next v

        ' Whenever A2 changes, clear A8 and reset its formatting.
        ' This prevents old or invalid values from being left behind.
        With Me.Range("A8")
            .ClearContents
            .Font.Italic = False
            .Font.Color = vbBlack
        End With

        ' If A2 contains one of the trigger values...
        If trigger Then

            ' Remove any existing validation rules from A8.
            Me.Range("A8").Validation.Delete

            ' Apply a dropdown list with exactly two allowed options.
            Me.Range("A8").Validation.Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="Option A,Option B"   ' <-- Replace with your two allowed choices

            ' Insert a placeholder message to guide the user.
            ' The placeholder is red and italic to clearly indicate
            ' that it is an instruction, not a real value.
            With Me.Range("A8")
                .Value = "Please select one of the available options"
                .Font.Italic = True
                .Font.Color = vbRed
            End With

        Else
            ' If A2 does NOT contain a trigger value,
            ' remove validation so A8 becomes a normal free‑text cell.
            Me.Range("A8").Validation.Delete
        End If

        ' Stop here so the A8 logic below does not run unnecessarily.
        Exit Sub
    End If

    ' =========================================================
    ' PART 2: What happens when cell A8 is changed
    ' =========================================================
    If Not Intersect(Target, Me.Range("A8")) Is Nothing Then

        ' When the user selects a real value from the dropdown,
        ' remove italics but keep the text red.
        ' This visually distinguishes real data from placeholder text.
        With Me.Range("A8")
            If .Value <> "" Then
                .Font.Italic = False
                .Font.Color = vbRed
            End If
        End With

    End If

End Sub


' =========================================================
' CLEANING FUNCTION TO FIX UNICODE / HYPHEN ISSUES
' =========================================================
Private Function CleanText(txt As String) As String
    ' This function removes or replaces hidden characters that can cause
    ' text comparisons to fail even when the text looks identical in Excel.

    Dim t As String
    t = txt

    ' Replace common Unicode dash characters with a normal hyphen
    t = Replace(t, ChrW(8211), "-")   ' en dash
    t = Replace(t, ChrW(8212), "-")   ' em dash
    t = Replace(t, ChrW(8209), "-")   ' non-breaking hyphen
    t = Replace(t, ChrW(173), "-")    ' soft hyphen

    ' Replace non-breaking spaces with normal spaces
    t = Replace(t, Chr(160), " ")

    ' Trim leading/trailing spaces
    t = Trim(t)

    CleanText = t
End Function
