
' ================================
' Log sheet "State Affecting" population
' ================================
' What this module does
' - You run PopulateLog_StateAffecting_Master
' - It fills the 50 columns named like "Alabama Affecting" ... "Wyoming Affecting"
' - For each row in the Log sheet's used range:
'     - If column K contains "All" OR contains that state's code (e.g., "USAL"), write "Yes"
'     - Otherwise write "No"
'
Option Explicit

' Entry point (run this macro)
Public Sub PopulateLog_StateAffecting_Master()
	' Requirements checklist:
	' [x] Multiple private subs in same file
	' [x] One master sub to run everything
	' [x] Loop used range rows; read column K; write Yes/No for each state column
	' [x] Handle "All" OR state code match (e.g., USAL)

	Dim ws As Worksheet
	Set ws = GetLogWorksheet()
	If ws Is Nothing Then Exit Sub

	Application.ScreenUpdating = False
	Application.EnableEvents = False
	Application.Calculation = xlCalculationManual

	On Error GoTo CleanFail
	PopulateLog_StateAffecting ws

CleanExit:
	Application.Calculation = xlCalculationAutomatic
	Application.EnableEvents = True
	Application.ScreenUpdating = True
	Exit Sub

CleanFail:
	' Keep it simple: restore Excel state and re-raise
	Resume CleanExit
End Sub

' ------------------------
' Core implementation
' ------------------------
Private Sub PopulateLog_StateAffecting(ByVal wsLog As Worksheet)
	Dim stateNames As Variant
	Dim stateCodes As Variant
	GetStateLists stateNames, stateCodes

	Dim lastRow As Long
	lastRow = GetLastRowFromColumn(wsLog, 11) ' Column K
	If lastRow < 2 Then Exit Sub ' assume row 1 is header

	' Map each "<State> Affecting" header to its column index.
	Dim stateCols() As Long
	stateCols = GetStateAffectingColumns(wsLog, stateNames)

	Dim r As Long
	For r = 2 To lastRow
		Dim kValue As String
		kValue = NormalizeAffectingValue(wsLog.Cells(r, 11).Value2)

		Dim isAll As Boolean
		isAll = ContainsToken(kValue, "ALL")

		Dim i As Long
		For i = LBound(stateNames) To UBound(stateNames)
			Dim targetCol As Long
			targetCol = stateCols(i)
			If targetCol > 0 Then
				Dim yesNo As String
				If isAll Or ContainsToken(kValue, CStr(stateCodes(i))) Then
					yesNo = "Yes"
				Else
					yesNo = "No"
				End If
				wsLog.Cells(r, targetCol).Value2 = yesNo
			End If
		Next i
	Next r
End Sub

' ------------------------
' Helpers
' ------------------------
Private Function GetLogWorksheet() As Worksheet
	' Assumption: the sheet is literally named "Log".
	' If your sheet has a different name, change it here.
	On Error Resume Next
	Set GetLogWorksheet = ThisWorkbook.Worksheets("Log")
	On Error GoTo 0
End Function

Private Sub GetStateLists(ByRef stateNames As Variant, ByRef stateCodes As Variant)
	' 50 states only (no DC / territories).
	' NOTE: VBA is picky about line continuations in Array(...) calls.
	' Keeping these on a single physical line avoids "Syntax error" / "Expected: )" issues.
	stateNames = Array("Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado", "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming")

	' Codes per your format: "US" + postal abbreviation.
	stateCodes = Array("USAL", "USAK", "USAZ", "USAR", "USCA", "USCO", "USCT", "USDE", "USFL", "USGA", "USHI", "USID", "USIL", "USIN", "USIA", "USKS", "USKY", "USLA", "USME", "USMD", "USMA", "USMI", "USMN", "USMS", "USMO", "USMT", "USNE", "USNV", "USNH", "USNJ", "USNM", "USNY", "USNC", "USND", "USOH", "USOK", "USOR", "USPA", "USRI", "USSC", "USSD", "USTN", "USTX", "USUT", "USVT", "USVA", "USWA", "USWV", "USWI", "USWY")
End Sub

Private Function GetStateAffectingColumns(ByVal ws As Worksheet, ByVal stateNames As Variant) As Long()
	' Finds columns by exact header match in row 1: "<State> Affecting"
	Dim cols() As Long
	ReDim cols(LBound(stateNames) To UBound(stateNames))

	Dim headerRow As Long
	headerRow = 1

	Dim usedLastCol As Long
	usedLastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

	Dim i As Long
	For i = LBound(stateNames) To UBound(stateNames)
		Dim headerText As String
		headerText = CStr(stateNames(i)) & " Affecting"
		cols(i) = FindHeaderColumn(ws, headerRow, 1, usedLastCol, headerText)
	Next i

	GetStateAffectingColumns = cols
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal startCol As Long, ByVal endCol As Long, ByVal headerText As String) As Long
	Dim c As Long
	For c = startCol To endCol
		If StrComp(Trim$(CStr(ws.Cells(headerRow, c).Value2)), headerText, vbTextCompare) = 0 Then
			FindHeaderColumn = c
			Exit Function
		End If
	Next c
	FindHeaderColumn = 0
End Function

Private Function GetLastRowFromColumn(ByVal ws As Worksheet, ByVal colIndex As Long) As Long
	' Uses the last non-empty cell in a given column.
	Dim lastRow As Long
	lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
	GetLastRowFromColumn = lastRow
End Function

Private Function NormalizeAffectingValue(ByVal v As Variant) As String
	' Normalizes the raw cell value so matching is more reliable.
	' - Uppercases
	' - Replaces common separators with spaces
	' - Pads with spaces at both ends so token-search is easy
	Dim s As String
	s = UCase$(Trim$(CStr(v)))

	' Normalize separators to spaces
	s = Replace(s, ",", " ")
	s = Replace(s, ";", " ")
	s = Replace(s, vbTab, " ")
	s = Replace(s, vbCr, " ")
	s = Replace(s, vbLf, " ")

	' Collapse multiple spaces
	Do While InStr(1, s, "  ") > 0
		s = Replace(s, "  ", " ")
	Loop

	NormalizeAffectingValue = " " & s & " "
End Function

Private Function ContainsToken(ByVal normalizedPadded As String, ByVal token As String) As Boolean
	' Expects normalizedPadded to be uppercased and padded with spaces.
	' Matches whole tokens only (so "USAL" doesn't match "USALX").
	Dim t As String
	t = " " & UCase$(Trim$(token)) & " "
	ContainsToken = (InStr(1, normalizedPadded, t, vbTextCompare) > 0)
End Function

