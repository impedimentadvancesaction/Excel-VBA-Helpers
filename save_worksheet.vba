Public Sub ExportNotificationClean()

    Dim newWb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim exportName As String
    Dim exportPath As String
    Dim fullPath As String

    ' Get filename from A1
    exportName = Trim(CStr(ThisWorkbook.Worksheets("Draft").Range("A1").Value))
    If Len(exportName) = 0 Then
        MsgBox "Cell A1 is empty — cannot create filename.", vbExclamation
        Exit Sub
    End If

    ' Remove invalid filename characters
    exportName = Replace(exportName, "/", "-")
    exportName = Replace(exportName, "\", "-")
    exportName = Replace(exportName, ":", "-")
    exportName = Replace(exportName, "*", "-")
    exportName = Replace(exportName, "?", "-")
    exportName = Replace(exportName, """", "'")
    exportName = Replace(exportName, "<", "(")
    exportName = Replace(exportName, ">", ")")
    exportName = Replace(exportName, "|", "-")

    ' Build full path in same folder as macro workbook
    exportPath = ThisWorkbook.Path
    fullPath = exportPath & "\" & exportName & ".xlsx"

    ' Copy the sheet to a new workbook
    ThisWorkbook.Worksheets("Draft").Copy
    Set newWb = ActiveWorkbook
    Set ws = newWb.Sheets(1)

    ' Remove all buttons (Form Controls and ActiveX)
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Or shp.Type = msoOLEControlObject Then
            shp.Delete
        End If
    Next shp

    ' Save as clean XLSX
    newWb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close False

    MsgBox "Notification exported as:" & vbCrLf & fullPath, vbInformation

End Sub
