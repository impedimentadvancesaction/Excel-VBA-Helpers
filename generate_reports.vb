
Option Explicit

' ================================
' Report generation (Tab1/Tab2/Tab3)
' ================================
' Assumptions (edit these names if your workbook uses different sheet names):
' - "tab1" = data source
' - "tab2" = report template (must remain unchanged)
' - "tab3" = lookup table
'
' What this macro does:
' 1) Sanitises Tab1 columns A, J, K, L (trim, normalize NBSP, remove line breaks)
' 2) Replaces Tab1 column J values when they fully match Tab3 column A (regex-anchored full match)
'    using the replacement from Tab3 column C
' 3) For each unique value in Tab1 column A, creates a new workbook from the Tab2 template,
'    writes:
'      - B1 = Tab1 column A
'      - B2 = Tab1 column J + newline + "No issues"
'      - B3 = Tab1 column K
'      - B4 = Tab1 column L
'    then saves both .xlsx and .pdf into the current workbook folder.
Public Sub GenerateReports_FromTab1TemplateLookup()
	Const DATA_SHEET_NAME As String = "Insert Info Here"
	Const TEMPLATE_SHEET_NAME As String = "Release Note"
	Const LOOKUP_SHEET_NAME As String = "PRJ CODES"

	Const DATA_COL_A As Long = 1  ' A
	Const DATA_COL_J As Long = 10 ' J
	Const DATA_COL_K As Long = 11 ' K
	Const DATA_COL_L As Long = 12 ' L

	Dim wsData As Worksheet, wsTemplate As Worksheet, wsLookup As Worksheet
	Set wsData = GetWorksheetByName(DATA_SHEET_NAME)
	Set wsTemplate = GetWorksheetByName(TEMPLATE_SHEET_NAME)
	Set wsLookup = GetWorksheetByName(LOOKUP_SHEET_NAME)
	If wsData Is Nothing Or wsTemplate Is Nothing Or wsLookup Is Nothing Then
		MsgBox "Missing one or more required sheets: " & DATA_SHEET_NAME & ", " & TEMPLATE_SHEET_NAME & ", " & LOOKUP_SHEET_NAME, vbExclamation, "Generate Reports"
		Exit Sub
	End If

	Dim lastRow As Long
	lastRow = GetMaxLastRow(wsData, Array(DATA_COL_A, DATA_COL_J, DATA_COL_K, DATA_COL_L))
	If lastRow < 2 Then
		MsgBox "No data rows found on " & DATA_SHEET_NAME & ".", vbInformation, "Generate Reports"
		Exit Sub
	End If

	Dim oldScreenUpdating As Boolean, oldEnableEvents As Boolean, oldDisplayAlerts As Boolean
	Dim oldCalculation As XlCalculation
	oldScreenUpdating = Application.ScreenUpdating
	oldEnableEvents = Application.EnableEvents
	oldDisplayAlerts = Application.DisplayAlerts
	oldCalculation = Application.Calculation

	Application.ScreenUpdating = False
	Application.EnableEvents = False
	Application.DisplayAlerts = False
	Application.Calculation = xlCalculationManual

	On Error GoTo CleanFail

	' 1) Sanitize Tab1 columns A/J/K/L
	SanitizeSheetColumns wsData, 2, lastRow, Array(DATA_COL_A, DATA_COL_J, DATA_COL_K, DATA_COL_L)

	' 2) Build lookup and apply exact replacements to Tab1 column J
	Dim lookupDict As Object
	Set lookupDict = BuildLookupDictionary(wsLookup) ' key: Tab3 col A, value: Tab3 col C
	ApplyLookupReplacements_ColumnJ wsData, 2, lastRow, DATA_COL_J, lookupDict

	' 3) One report per unique value in Tab1 column A
	Dim uniqueKeys As Object
	Set uniqueKeys = CreateObject("Scripting.Dictionary")
	uniqueKeys.CompareMode = 1 ' vbTextCompare

	Dim r As Long
	For r = 2 To lastRow
		Dim keyA As String
		keyA = Trim$(CStr(wsData.Cells(r, DATA_COL_A).Value2))
		If Len(keyA) > 0 Then
			If Not uniqueKeys.Exists(keyA) Then uniqueKeys.Add keyA, r ' store first row for that key
		End If
	Next r

	If uniqueKeys.Count = 0 Then
		MsgBox "No unique values found in " & DATA_SHEET_NAME & " column A.", vbInformation, "Generate Reports"
		GoTo CleanExit
	End If

	Dim baseFolder As String
	baseFolder = ThisWorkbook.Path
	If Len(baseFolder) = 0 Then baseFolder = CurDir$

	Dim createdCount As Long
	createdCount = 0

	Dim k As Variant
	For Each k In uniqueKeys.Keys
		Dim sourceRow As Long
		sourceRow = CLng(uniqueKeys(k))

		Dim vA As String, vJ As String, vK As String, vL As String
		vA = CStr(wsData.Cells(sourceRow, DATA_COL_A).Value2)
		vJ = CStr(wsData.Cells(sourceRow, DATA_COL_J).Value2)
		vK = CStr(wsData.Cells(sourceRow, DATA_COL_K).Value2)
		vL = CStr(wsData.Cells(sourceRow, DATA_COL_L).Value2)

		CreateOneReportFromTemplate wsTemplate, baseFolder, vA, vJ, vK, vL
		createdCount = createdCount + 1
	Next k

	MsgBox "Created " & createdCount & " report(s) in: " & baseFolder, vbInformation, "Generate Reports"

CleanExit:
	Application.Calculation = oldCalculation
	Application.DisplayAlerts = oldDisplayAlerts
	Application.EnableEvents = oldEnableEvents
	Application.ScreenUpdating = oldScreenUpdating
	Exit Sub

CleanFail:
	' Best effort restore Excel state
	Resume CleanExit
End Sub

' ------------------------
' Step 1: Sanitise cells
' ------------------------
Private Sub SanitizeSheetColumns(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal cols As Variant)
	If ws Is Nothing Then Exit Sub
	If lastRow < firstRow Then Exit Sub

	Dim i As Long
	For i = LBound(cols) To UBound(cols)
		Dim colIndex As Long
		colIndex = CLng(cols(i))

		Dim r As Long
		For r = firstRow To lastRow
			Dim v As Variant
			v = ws.Cells(r, colIndex).Value2
			If Not IsError(v) Then
				If VarType(v) = vbString Then
					ws.Cells(r, colIndex).Value2 = SanitizeText(CStr(v))
				End If
			End If
		Next r
	Next i
End Sub

Private Function SanitizeText(ByVal s As String) As String
	Dim t As String
	' Convert non-breaking spaces to normal spaces
	t = Replace(s, ChrW$(160), " ")
	' Replace any line breaks with spaces
	t = Replace(t, vbCrLf, " ")
	t = Replace(t, vbCr, " ")
	t = Replace(t, vbLf, " ")
	SanitizeText = Trim$(t)
End Function

' ------------------------
' Step 2: Lookup replace (Tab3 A -> C) on Tab1 column J
' ------------------------
Private Function BuildLookupDictionary(ByVal wsLookup As Worksheet) As Object
	Dim dict As Object
	Set dict = CreateObject("Scripting.Dictionary")
	dict.CompareMode = 1 ' vbTextCompare

	If wsLookup Is Nothing Then
		Set BuildLookupDictionary = dict
		Exit Function
	End If

	Dim lastRow As Long
	lastRow = wsLookup.Cells(wsLookup.Rows.Count, 1).End(xlUp).Row ' column A
	If lastRow < 1 Then
		Set BuildLookupDictionary = dict
		Exit Function
	End If

	Dim r As Long
	For r = 1 To lastRow
		Dim fromValue As String
		fromValue = SanitizeText(CStr(wsLookup.Cells(r, 1).Value2)) ' A
		If Len(fromValue) > 0 Then
			Dim toValue As String
			toValue = SanitizeText(CStr(wsLookup.Cells(r, 3).Value2)) ' C
			If Not dict.Exists(fromValue) Then
				dict.Add fromValue, toValue
			End If
		End If
	Next r

	Set BuildLookupDictionary = dict
End Function

Private Sub ApplyLookupReplacements_ColumnJ(ByVal wsData As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal colJ As Long, ByVal lookupDict As Object)
	If wsData Is Nothing Then Exit Sub
	If lookupDict Is Nothing Then Exit Sub
	If lastRow < firstRow Then Exit Sub

	' Late-bound regex to avoid requiring a reference
	Dim re As Object
	Set re = CreateObject("VBScript.RegExp")
	re.Global = False
	re.IgnoreCase = True

	Dim r As Long
	For r = firstRow To lastRow
		Dim cellValue As Variant
		cellValue = wsData.Cells(r, colJ).Value2
		If Not IsError(cellValue) Then
			Dim s As String
			s = SanitizeText(CStr(cellValue))
			If Len(s) > 0 Then
				If lookupDict.Exists(s) Then
					' Full match only (anchored)
					re.Pattern = "^" & EscapeRegex(s) & "$"
					If re.Test(s) Then
						wsData.Cells(r, colJ).Value2 = CStr(lookupDict(s))
					End If
				End If
			Else
				' Keep truly blank cells blank (don’t write back whitespace)
				If Len(CStr(cellValue)) > 0 Then wsData.Cells(r, colJ).Value2 = vbNullString
			End If
		End If
	Next r
End Sub

Private Function EscapeRegex(ByVal s As String) As String
	' Escapes regex metacharacters for VBScript.RegExp
	Dim specials As String
	specials = "\\.^$|?*+()[]{}" ' backslash must be doubled in VBA string

	Dim i As Long
	Dim ch As String
	Dim result As String

	For i = 1 To Len(s)
		ch = Mid$(s, i, 1)
		If InStr(1, specials, ch, vbBinaryCompare) > 0 Then
			result = result & "\" & ch
		Else
			result = result & ch
		End If
	Next i

	EscapeRegex = result
End Function

' ------------------------
' Step 3: Create reports without altering template
' ------------------------
Private Sub CreateOneReportFromTemplate(ByVal wsTemplate As Worksheet, ByVal baseFolder As String, ByVal vA As String, ByVal vJ As String, ByVal vK As String, ByVal vL As String)
	If wsTemplate Is Nothing Then Exit Sub

	' Copy template sheet to a brand-new workbook (keeps the original template untouched)
	wsTemplate.Copy

	Dim wbReport As Workbook
	Set wbReport = ActiveWorkbook
	Dim wsReport As Worksheet
	Set wsReport = wbReport.Worksheets(1)

	wsReport.Range("B1").Value2 = vA
	wsReport.Range("B2").Value2 = vJ & vbCrLf & "No issues"
	wsReport.Range("B2").WrapText = True
	wsReport.Range("B3").Value2 = vK
	wsReport.Range("B4").Value2 = vL

	Dim baseName As String
	baseName = SafeFileName(CStr(wsReport.Range("B1").Value2))
	If Len(baseName) = 0 Then baseName = "Report"

	Dim xlsxPath As String
	xlsxPath = MakeUniqueFilePath(baseFolder, baseName, "xlsx")

	Dim pdfPath As String
	pdfPath = MakeUniqueFilePath(baseFolder, baseName, "pdf")

	' Save workbook
	wbReport.SaveAs Filename:=xlsxPath, FileFormat:=xlOpenXMLWorkbook
	' Export PDF
	wbReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

	' Close the generated workbook
	wbReport.Close SaveChanges:=False
End Sub

Private Function SafeFileName(ByVal s As String) As String
	Dim t As String
	t = Trim$(CStr(s))

	' Replace invalid filename chars on Windows
	Dim badChars As Variant
	badChars = Array("\\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)

	Dim i As Long
	For i = LBound(badChars) To UBound(badChars)
		t = Replace(t, CStr(badChars(i)), "_")
	Next i

	' Avoid trailing dots/spaces (Windows restriction)
	Do While Len(t) > 0 And (Right$(t, 1) = "." Or Right$(t, 1) = " ")
		t = Left$(t, Len(t) - 1)
	Loop

	' Reasonable length limit
	If Len(t) > 150 Then t = Left$(t, 150)

	SafeFileName = t
End Function

Private Function MakeUniqueFilePath(ByVal folderPath As String, ByVal baseName As String, ByVal extensionNoDot As String) As String
	Dim folder As String
	folder = folderPath
	If Len(folder) = 0 Then folder = CurDir$
	If Right$(folder, 1) <> "\\" Then folder = folder & "\\"

	Dim ext As String
	ext = "." & extensionNoDot

	Dim candidate As String
	candidate = folder & baseName & ext
	If Len(Dir$(candidate)) = 0 Then
		MakeUniqueFilePath = candidate
		Exit Function
	End If

	Dim i As Long
	For i = 2 To 999
		candidate = folder & baseName & " (" & CStr(i) & ")" & ext
		If Len(Dir$(candidate)) = 0 Then
			MakeUniqueFilePath = candidate
			Exit Function
		End If
	Next i

	' Fallback
	MakeUniqueFilePath = folder & baseName & " (" & Format$(Now, "yyyymmdd_hhnnss") & ")" & ext
End Function

' ------------------------
' Generic helpers
' ------------------------
Private Function GetWorksheetByName(ByVal sheetName As String) As Worksheet
	On Error Resume Next
	Set GetWorksheetByName = ThisWorkbook.Worksheets(sheetName)
	On Error GoTo 0
End Function

Private Function GetMaxLastRow(ByVal ws As Worksheet, ByVal colIndices As Variant) As Long
	Dim maxRow As Long
	maxRow = 0

	Dim i As Long
	For i = LBound(colIndices) To UBound(colIndices)
		Dim c As Long
		c = CLng(colIndices(i))
		Dim lr As Long
		lr = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
		If lr > maxRow Then maxRow = lr
	Next i

	GetMaxLastRow = maxRow
End Function

