Option Explicit

' =============================================================================
' ConsolidateColumnAFromFolder
' Opens each Excel file in a folder, reads the value from cell A4 of each.
' If A4 contains multiple lines (Alt+Enter), each line is written to its own
' row: column A = line text, column B = source filename.
' =============================================================================
Public Sub ConsolidateColumnAFromFolder()
    Const FOLDER_PATH As String = "C:\YourFolder"   ' Change this or use FolderPicker below
    Dim folderPath As String
    Dim wbDest As Workbook
    Dim wsDest As Worksheet
    Dim lastRowDest As Long
    Dim destRow As Long
    Dim f As String
    Dim calcMode As Long
    Dim eventsEnabled As Boolean
    Dim extensions As Variant
    Dim e As Long
    Dim totalFiles As Long
    Dim fileNum As Long

    Set wbDest = ThisWorkbook
    Set wsDest = wbDest.ActiveSheet

    ' Optional: use folder picker instead of constant (uncomment to use)
    ' folderPath = GetFolderPath()
    ' If folderPath = "" Then Exit Sub
    folderPath = FOLDER_PATH
    If Right(folderPath, 1) = "\" Then folderPath = Left(folderPath, Len(folderPath) - 1)

    ' Performance: turn off screen updates, calculation, and events
    Application.ScreenUpdating = False
    calcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    eventsEnabled = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Cleanup

    ' Count files first for progress indicator (excludes current workbook)
    totalFiles = CountExcelFilesInFolder(folderPath, wbDest)

    ' Find next empty row in column A on destination sheet
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row
    destRow = IIf(lastRowDest = 1 And Len(Trim(wsDest.Cells(1, 1).Value)) = 0, 1, lastRowDest + 1)

    ' Process common Excel file types
    fileNum = 0
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")
    For e = LBound(extensions) To UBound(extensions)
        f = Dir(folderPath & "\" & extensions(e))
        Do While f <> ""
            If StrComp(folderPath & "\" & f, wbDest.Path & "\" & wbDest.Name, vbTextCompare) <> 0 Then
                fileNum = fileNum + 1
                Application.StatusBar = "Processing file " & fileNum & " of " & totalFiles & "..."
                ConsolidateOneFile folderPath, f, wsDest, destRow
            End If
            f = Dir()
        Loop
    Next e

Cleanup:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    Application.EnableEvents = eventsEnabled
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' Returns the number of Excel files in folder (excluding wbDest if it is in that folder).
Private Function CountExcelFilesInFolder(ByVal folderPath As String, ByVal wbDest As Workbook) As Long
    Dim f As String
    Dim extensions As Variant
    Dim e As Long
    CountExcelFilesInFolder = 0
    extensions = Array("*.xlsx", "*.xlsm", "*.xls")
    For e = LBound(extensions) To UBound(extensions)
        f = Dir(folderPath & "\" & extensions(e))
        Do While f <> ""
            If StrComp(folderPath & "\" & f, wbDest.Path & "\" & wbDest.Name, vbTextCompare) <> 0 Then
                CountExcelFilesInFolder = CountExcelFilesInFolder + 1
            End If
            f = Dir()
        Loop
    Next e
End Function

' Opens one workbook, reads cell A4 from the first sheet, splits by line breaks,
' and appends one row per line to wsDest (A = line, B = filename).
Private Sub ConsolidateOneFile(ByVal folderPath As String, ByVal fileName As String, _
    ByVal wsDest As Worksheet, ByRef destRow As Long)
    Dim wbSrc As Workbook
    Dim fullPath As String
    Dim cellValue As String
    Dim lines As Variant
    Dim i As Long

    fullPath = folderPath & "\" & fileName
    On Error Resume Next
    Set wbSrc = Workbooks.Open(fullPath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    On Error GoTo 0
    If wbSrc Is Nothing Then Exit Sub

    On Error GoTo CloseSrc
    ' Get A4 as string (handles numbers/dates) and normalize line breaks (CRLF/CR -> LF)
    cellValue = CStr(wbSrc.Sheets(1).Range("A4").Value)
    cellValue = Replace(cellValue, vbCrLf, vbLf)
    cellValue = Replace(cellValue, vbCr, vbLf)
    lines = Split(cellValue, vbLf)

    For i = LBound(lines) To UBound(lines)
        wsDest.Cells(destRow, 1).Value = lines(i)
        wsDest.Cells(destRow, 2).Value = fileName
        destRow = destRow + 1
    Next i

CloseSrc:
    wbSrc.Close SaveChanges:=False
End Sub

' Optional: returns a folder path from the user, or empty string if cancelled.
Public Function GetFolderPath() As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Select folder containing workbooks to consolidate"
    If dlg.Show = -1 Then GetFolderPath = dlg.SelectedItems(1)
End Function
