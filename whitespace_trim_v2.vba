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

    ' Prevent recursive triggering while we modify cell values
    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' Load the changed range into memory for fast processing
    data = Target.Value

    ' Iterate through every cell in the changed range
    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            data(r, c) = CleanCellValue(data(r, c))
        Next c
    Next r

    ' Write the cleaned values back to the worksheet
    Target.Value = data

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
    Dim i As Long

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

    Next i

    ' Reassemble the original multi-line structure
    CleanCellValue = Join(lines, vbLf)

End Function
