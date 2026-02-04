Option Explicit

' Runner macro (shows up in the Macros dialog).
' Adjust these parameters to match your sheet/column.
Public Sub RemoveExtensions_Data_ColumnB()
	' Example run: Column B, starting at row 3, on the "Data" sheet in this workbook.
	RemoveFileExtensionsInColumn "B", 3, ThisWorkbook.Worksheets("Data")
End Sub

Private Sub RemoveFileExtensionsInColumn(
	Optional ByVal columnLetter As String = "A", 
	Optional ByVal firstDataRow As Long = 2, 
	Optional ByVal targetSheet As Worksheet)

	' Removes the last file extension (e.g. ".xlsx") from each cell in a column.
	'
	' Examples:
	'   "C:\Reports\file.xlsx"  -> "C:\Reports\file"
	'   "archive.tar.gz"         -> "archive.tar"
	'   "noextension"            -> unchanged
	'   ".gitignore"             -> unchanged

	' Decide which worksheet to operate on.
	Dim ws As Worksheet
	If targetSheet Is Nothing Then
		Set ws = ActiveSheet
	Else
		Set ws = targetSheet
	End If

	' Find the last used row in the target column.
	Dim lastRow As Long
	lastRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
	If lastRow < firstDataRow Then Exit Sub

	' Build a contiguous range we can read/write in one go (fast).
	Dim rng As Range
	Set rng = ws.Range(ws.Cells(firstDataRow, columnLetter), ws.Cells(lastRow, columnLetter))

	' Remember current Excel settings so we can restore them.
	Dim prevScreenUpdating As Boolean
	Dim prevEnableEvents As Boolean
	prevScreenUpdating = Application.ScreenUpdating
	prevEnableEvents = Application.EnableEvents

	' If anything errors, jump to CleanUp to restore settings.
	On Error GoTo CleanUp
	Application.ScreenUpdating = False
	Application.EnableEvents = False

	' Pull all values into memory (usually a 2D array).
	Dim data As Variant
	data = rng.Value2

	' rng.Value2 is usually a 2D array, but can be a scalar if rng is a single cell.
	If Not IsArray(data) Then
		Dim singleCell As Variant
		singleCell = data
		ReDim data(1 To 1, 1 To 1)
		data(1, 1) = singleCell
	End If

	Dim i As Long
	For i = LBound(data, 1) To UBound(data, 1)
		' Skip Excel error values (#N/A, #VALUE!, etc.).
		If Not IsError(data(i, 1)) Then
			' Skip blanks; otherwise strip the last extension.
			If Len(Trim$(CStr(data(i, 1)))) > 0 Then
				data(i, 1) = StripFileExtension(CStr(data(i, 1)))
			End If
		End If
	Next i

	' Write the updated values back to the sheet in one operation.
	rng.Value2 = data

CleanUp:
	Application.EnableEvents = prevEnableEvents
	Application.ScreenUpdating = prevScreenUpdating
End Sub

Private Function StripFileExtension(ByVal inputText As String) As String
	' Returns the input text with the last ".ext" removed, when it looks like a file extension.
	' Safe for full paths and for dotfiles like ".gitignore".
	Dim s As String
	s = Trim$(inputText)
	If Len(s) = 0 Then
		StripFileExtension = inputText
		Exit Function
	End If

	' Find the last path separator so we only consider dots in the filename portion.
	Dim lastSlash As Long
	Dim lastFwdSlash As Long
	lastSlash = InStrRev(s, "\")
	lastFwdSlash = InStrRev(s, "/")
	If lastFwdSlash > lastSlash Then lastSlash = lastFwdSlash

	' Find the last dot (.) which usually starts the extension.
	Dim lastDot As Long
	lastDot = InStrRev(s, ".")

	' Keep unchanged when:
	' - no dot
	' - dot is first character of filename (e.g. ".gitignore")
	' - dot is the last character
	If lastDot = 0 Then
		StripFileExtension = inputText
		Exit Function
	End If
	If lastDot <= lastSlash + 1 Then
		StripFileExtension = inputText
		Exit Function
	End If
	If lastDot = Len(s) Then
		StripFileExtension = inputText
		Exit Function
	End If

	StripFileExtension = Left$(s, lastDot - 1)
End Function

