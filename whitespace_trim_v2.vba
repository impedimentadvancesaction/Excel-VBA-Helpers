Option Explicit

'=========================================================
' Worksheet_Change
'---------------------------------------------------------
' Fires automatically whenever data is changed on the sheet.
' This includes paste operations of any size or shape.
'
' The routine:
'   • Processes only the cells that were changed
'   • Cleans text values while preserving structure
'   • Leaves numbers, dates, and formulas untouched
'=========================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim data As Variant
    Dim r As Long, c As Long
    Dim rngText As Range
    Dim area As Range

    ' Prevent recursive triggering while we modify cell values
    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' Only process text constants (ignores numbers, dates, blanks, and formulas)
    ' NOTE: When a single cell changes, Target.Value is a scalar (not a 2-D array).
    On Error Resume Next
    Set rngText = Target.SpecialCells(xlCellTypeConstants, xlTextValues)
    On Error GoTo SafeExit

    If rngText Is Nothing Then GoTo SafeExit

    ' Process each area to keep array writes fast and safe
    For Each area In rngText.Areas

        ' Load the changed range into memory for fast processing
        data = area.Value2

        If IsArray(data) Then
            ' Iterate through every cell in the changed range
            For r = 1 To UBound(data, 1)
                For c = 1 To UBound(data, 2)
                    data(r, c) = CleanCellValue(data(r, c))
                Next c
            Next r

            ' Write the cleaned values back to the worksheet
            area.Value2 = data
        Else
            ' Single-cell area (scalar)
            area.Value2 = CleanCellValue(data)
        End If

    Next area

SafeExit:
    ' Always re-enable events, even if an error occurs
    Application.EnableEvents = True

End Sub


'=========================================================
' CleanCellValue
'---------------------------------------------------------
' Cleans unwanted whitespace from a single cell value.
'
' Rules:
'   • Only text values are modified
'   • Leading and trailing spaces are removed
'   • Tabs and non-breaking spaces are normalised
'   • Multiple spaces collapse to a single space
'   • Line breaks are preserved (multi-line cells remain intact)
'
' This ensures data cleanliness without destroying meaning.
'=========================================================
Private Function CleanCellValue(ByVal v As Variant) As Variant

    Dim lines() As String
    Dim kept() As String
    Dim i As Long
    Dim k As Long

    ' Exit immediately if the value is not text
    If VarType(v) <> vbString Then
        CleanCellValue = v
        Exit Function
    End If

    ' Normalise Windows line endings to a single format
    v = Replace(v, vbCrLf, vbLf)

    ' Split the text into individual lines
    lines = Split(v, vbLf)

    ' Process each line independently
    For i = LBound(lines) To UBound(lines)

        ' Replace tabs with spaces
        lines(i) = Replace(lines(i), vbTab, " ")

        ' Replace non-breaking spaces (often from web data)
        lines(i) = Replace(lines(i), Chr(160), " ")

        ' Remove leading and trailing whitespace
        lines(i) = Trim(lines(i))

        ' Collapse multiple spaces into a single space
        Do While InStr(lines(i), "  ") > 0
            lines(i) = Replace(lines(i), "  ", " ")
        Loop

        ' Remove empty lines (including those that were only whitespace)
        If Len(lines(i)) > 0 Then
            If k = 0 Then
                ReDim kept(0 To 0)
            Else
                ReDim Preserve kept(0 To k)
            End If
            kept(k) = lines(i)
            k = k + 1
        End If

    Next i

    ' Reassemble the original multi-line structure
    If k = 0 Then
        CleanCellValue = vbNullString
    Else
        CleanCellValue = Join(kept, vbLf)
    End If

End Function
