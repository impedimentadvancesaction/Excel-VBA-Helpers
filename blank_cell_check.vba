
Option Explicit

' Scans an external workbook (hard-coded path) for blank cells in used columns
' (A through the last used column on each sheet) from row 6 down to the last used row,
' and writes a report table to this workbook.
Public Sub ReportBlankCellsFromExternalFile()
	Const SOURCE_FILE_PATH As String = "C:\Path\To\ExternalWorkbook.xlsx" ' <- CHANGE THIS

	Dim srcWb As Workbook
	Dim srcWs As Worksheet
	Dim reportWs As Worksheet
	Dim reportRow As Long

	Application.ScreenUpdating = False
	Application.EnableEvents = False

	On Error GoTo CleanFail

	If Len(Dir$(SOURCE_FILE_PATH)) = 0 Then
		Err.Raise vbObjectError + 2000, , "Source file not found: " & SOURCE_FILE_PATH
	End If

	' Create/clear report sheet in the current workbook.
	Set reportWs = GetOrCreateWorksheet(ThisWorkbook, "Blank Cell Report")
	reportWs.Cells.Clear

	' Report headers.
	reportWs.Range("A1").Value = "Source File"
	reportWs.Range("B1").Value = "Sheet"
	reportWs.Range("C1").Value = "Row"
	reportWs.Range("D1").Value = "Column"
	reportWs.Range("E1").Value = "Column Letter"
	reportWs.Range("F1").Value = "Cell Address"
	reportWs.Range("G1").Value = "Cell Value"
	reportWs.Range("H1").Value = "Note"

	reportWs.Range("A1:H1").Font.Bold = True
	reportRow = 2

	' Open source workbook read-only.
	Set srcWb = Workbooks.Open(Filename:=SOURCE_FILE_PATH, ReadOnly:=True)

	Dim lastRow As Long
	Dim lastCol As Long
	Dim r As Long
	Dim c As Long
	Dim cell As Range
	Dim valueText As String

	For Each srcWs In srcWb.Worksheets
		lastCol = LastUsedColumn(srcWs)
		If lastCol < 1 Then
			GoTo NextSheet
		End If

		lastRow = LastUsedRowInColumns(srcWs, 1, lastCol)
		If lastRow < 6 Then
			' Nothing to scan on this sheet.
			GoTo NextSheet
		End If

		For r = 6 To lastRow
			For c = 1 To lastCol
				Set cell = srcWs.Cells(r, c)

				' Treat truly empty and formula-returning-empty as blank.
				valueText = CStr(cell.Value)
				If Len(Trim$(valueText)) = 0 Then
					reportWs.Cells(reportRow, 1).Value = SOURCE_FILE_PATH
					reportWs.Cells(reportRow, 2).Value = srcWs.Name
					reportWs.Cells(reportRow, 3).Value = r
					reportWs.Cells(reportRow, 4).Value = c
					reportWs.Cells(reportRow, 5).Value = ColumnLetterFromNumber(c)
					reportWs.Cells(reportRow, 6).Value = cell.Address(False, False)
					reportWs.Cells(reportRow, 7).Value = valueText
					reportWs.Cells(reportRow, 8).Value = IIf(cell.HasFormula, "Formula returns blank", "Empty")
					reportRow = reportRow + 1
				End If
			Next c
		Next r

NextSheet:
	Next srcWs

	' Format report as a table-like range.
	If reportRow > 2 Then
		With reportWs.Range("A1:H" & reportRow - 1)
			.Columns.AutoFit
			.Borders.LineStyle = xlContinuous
		End With
	Else
		reportWs.Range("A2").Value = "No blanks found in used columns (A to last used column per sheet) from row 6 down."
		reportWs.Columns.AutoFit
	End If

CleanExit:
	On Error Resume Next
	If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
	Application.EnableEvents = True
	Application.ScreenUpdating = True
	Exit Sub

CleanFail:
	MsgBox "Report failed: " & Err.Description, vbExclamation
	Resume CleanExit
End Sub

Private Function LastUsedRowInColumns(ws As Worksheet, firstCol As Long, lastCol As Long) As Long
	' Finds the last used row across a column span without being fooled by formatting.
	Dim searchRange As Range
	Set searchRange = ws.Range(ws.Cells(1, firstCol), ws.Cells(ws.Rows.Count, lastCol))

	Dim lastCell As Range
	Set lastCell = searchRange.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
									SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

	If lastCell Is Nothing Then
		LastUsedRowInColumns = 0
	Else
		LastUsedRowInColumns = lastCell.Row
	End If
End Function

Private Function LastUsedColumn(ws As Worksheet) As Long
	' Finds the last used column on a sheet without being fooled by formatting.
	Dim lastCell As Range
	Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
							SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

	If lastCell Is Nothing Then
		LastUsedColumn = 0
	Else
		LastUsedColumn = lastCell.Column
	End If
End Function

Private Function GetOrCreateWorksheet(wb As Workbook, sheetName As String) As Worksheet
	On Error GoTo Create
	Set GetOrCreateWorksheet = wb.Worksheets(sheetName)
	Exit Function
Create:
	On Error GoTo 0
	Set GetOrCreateWorksheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
	GetOrCreateWorksheet.Name = sheetName
End Function

Private Function ColumnLetterFromNumber(colNum As Long) As String
	ColumnLetterFromNumber = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function
