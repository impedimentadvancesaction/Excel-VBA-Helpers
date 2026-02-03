Option Explicit

Public Sub GenerateReports()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet, ws5 As Worksheet
    Dim dictCRQ As Object, dictLookup As Object
    Dim lastRow1 As Long, lastRow2 As Long, r As Long, crq As Variant
    
    Set ws1 = Worksheets("Working Sheet")
    Set ws2 = Worksheets("GCA Export")
    Set ws3 = Worksheets("R9-RFM-CRQ-New-Jersey-Regulator")
    Set ws4 = Worksheets("Game Configurations-Activations")
    Set ws5 = Worksheets("Tracking")  ' UPDATE THIS TO YOUR ACTUAL TRACKING SHEET NAME
    
    ' Snapshot templates
    ws3.Copy after:=ws4
    ActiveSheet.Name = "TMPREGFORM_1"
    ws4.Copy after:=Sheets("TMPREGFORM_1")
    ActiveSheet.Name = "TMPREGFORM_2"
    
    ' Collect unique CRQs
    Set dictCRQ = CreateObject("Scripting.Dictionary")
    lastRow1 = ws1.Cells(ws1.Rows.Count, 3).End(xlUp).Row
    For r = 2 To lastRow1
        crq = Trim(ws1.Cells(r, 3).Value)
        If Len(crq) > 0 Then If Not dictCRQ.Exists(crq) Then dictCRQ.Add crq, 1
    Next
    
    ' Build lookup dictionary from Tab 2
    Set dictLookup = CreateObject("Scripting.Dictionary")
    lastRow2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    Dim colRegID As Long: colRegID = ws2.Rows(1).Find("RegulatedGameID").Column
    For r = 2 To lastRow2
        dictLookup(Trim(ws2.Cells(r, colRegID).Value)) = r
    Next
    
    ' Loop through each CRQ
    Dim key As Variant
    For Each key In dictCRQ.Keys
        crq = key
        
        ' --- Get Provider for the CRQ from Tab 1 Column E ---
        Dim provider As String
        Dim providerND13 As String
        For r = 2 To lastRow1
            If ws1.Cells(r, 3).Value = crq Then
                provider = ws1.Cells(r, 5).Value
            End If
        Next
        
        ' Check if provider is Hillside Games - use "In House Games" for ND-13
        If provider = "Hillside Games" Then
            providerND13 = "In House Games"
        Else
            providerND13 = provider
        End If
        
        ws4.Columns(7).NumberFormat = "0"
        
        ' --- Populate Tab 4 ---
        Dim outRow As Long: outRow = 3
        Dim rgID As Variant, rr As Long
        For r = 2 To lastRow1
            If Trim(ws1.Cells(r, 3).Value) = crq Then
                rgID = ws1.Cells(r, 4).Value
                If dictLookup.Exists(CStr(rgID)) Then
                    rr = dictLookup(CStr(rgID))
                    ws4.Cells(outRow, 1).Value = crq
                    ws4.Cells(outRow, 2).Value = ws2.Cells(rr, ws2.Rows(1).Find("GameName").Column).Value
                    ws4.Cells(outRow, 3).Value = provider
                    ws4.Cells(outRow, 4).Value = ws2.Cells(rr, ws2.Rows(1).Find("GamesCertificateReference").Column).Value
                    ws4.Cells(outRow, 6).Value = ws2.Cells(rr, ws2.Rows(1).Find("Delivery Channel").Column).Value
                    ws4.Cells(outRow, 7).Value = ws2.Cells(rr, ws2.Rows(1).Find("GameVersion").Column).Value
                    ws4.Cells(outRow, 8).Value = ws2.Cells(rr, ws2.Rows(1).Find("TheoreticalRTP").Column).Value
                    ws4.Cells(outRow, 9).Value = IIf(LCase(Trim(ws2.Cells(rr, ws2.Rows(1).Find("isProgressive").Column).Value)) = "true", "Y", "N")
                    ws4.Cells(outRow, 10).Value = IIf(LCase(Trim(ws2.Cells(rr, ws2.Rows(1).Find("GameType").Column).Value)) = "slots", "Y", "N")
                    ws4.Cells(outRow, 11).Value = Format(NextBusinessDay(ParseDDMMYYYY(ws1.Cells(rr, ws1.Rows(1).Find("dd/mm/yyyy").Column).Value)), "dd/mm/yyyy")
                    outRow = outRow + 1
                End If
            End If
        Next
        
        ' --- Populate Tab 3 ---
        Dim uniqNames As Object: Set uniqNames = CreateObject("Scripting.Dictionary")
        Dim lastRow4 As Long: lastRow4 = ws4.Cells(ws4.Rows.Count, 2).End(xlUp).Row
        For r = 3 To lastRow4
            If Len(Trim(ws4.Cells(r, 2).Value)) > 0 Then uniqNames(Trim(ws4.Cells(r, 2).Value)) = 1
        Next
        
        If uniqNames.Count = 1 Then
            ws3.Range("A1").Value = "IG-ND-13(HR) Activate " & providerND13 & " game " & ws4.Cells(3, 2).Value & " on Desktop & Mobile (" & crq & ")"
            ws3.Range("A4").Value = "Activate " & providerND13 & " game " & ws4.Cells(3, 2).Value & " on Desktop & Mobile " & vbCrLf & "New Content"
        ElseIf uniqNames.Count > 1 Then
            ws3.Range("A1").Value = "IG-ND-13(HR) Activate multiple " & providerND13 & " games on Desktop & Mobile (" & crq & ")" & vbCrLf & "New Content"
            ws3.Range("A4").Value = "Activate multiple " & providerND13 & " games on Desktop & Mobile " & vbCrLf & "New Content"
        End If
        ws3.Range("A16").Value = Format(ParseDDMMYYYY(ws4.Range("K3").Value), "MMMM dd, yyyy") & " – 1 Hour"
        
        For r = 2 To lastRow1
            If ws1.Cells(r, 3).Value = crq Then
                ws3.Range("A19").Value = ws1.Cells(r, 2).Value
                Exit For
            End If
        Next
        ws3.Range("A28").Value = "See Games Installation Form " & crq
        ws3.Range("A34").Value = "See Games Installation Form " & crq
        
        ' Export ---
        With ws4.UsedRange
            .Borders.LineStyle = xlContinuous
        End With
        Dim lastRow4_2 As Long
        lastRow4_2 = ws4.Cells(ws4.Rows.Count, 1).End(xlUp).Row
        If lastRow4_2 >= 3 Then
            ws4.Range("E3:E" & lastRow4_2).Value = "Casino"
        End If
        With ws4.Range("A3:K" & lastRow4_2)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ExportSheet ws4, "Game Installation Form " & crq
        ws3.UsedRange.Font.Color = vbBlack
        ExportSheet ws3, ws3.Range("A1").Value
        
        ' --- Save tracking details to ws5 ---
        Dim nextRow5 As Long
        Dim ndFilename As String
        Dim jiraValue As String
        Dim gifDate As String
        Dim colJira As Long
        
        ' Find next available row in ws5
        nextRow5 = ws5.Cells(ws5.Rows.Count, 1).End(xlUp).Row + 1
        
        ' Get ND filename
        ndFilename = SanitizeFileName(ws3.Range("A1").Value) & ".xlsx"
        
        ' Get JIRA value from ws1
        colJira = ws1.Rows(1).Find("JIRA").Column
        For r = 2 To lastRow1
            If ws1.Cells(r, 3).Value = crq Then
                jiraValue = ws1.Cells(r, colJira).Value
                Exit For
            End If
        Next
        
        ' Get GIF date
        gifDate = ws4.Range("K3").Value
        
        ' Write to ws5
        ws5.Cells(nextRow5, 1).Value = crq                      ' Column A: CRQ
        ws5.Cells(nextRow5, 2).Value = ndFilename               ' Column B: ND filename
        ws5.Cells(nextRow5, 8).Value = jiraValue                ' Column H: JIRA value
        ws5.Cells(nextRow5, 11).Value = gifDate                 ' Column K: GIF date
        ws5.Cells(nextRow5, 12).Value = Date                    ' Column L: Today's date
        ws5.Cells(nextRow5, 14).Value = "Game Activation"       ' Column N: "Game Activation"
        
        ' --- Restore templates ---
        Application.DisplayAlerts = False
        Sheets("TMPREGFORM_1").Cells.Copy ws3.Cells
        Sheets("TMPREGFORM_2").Cells.Copy ws4.Cells
        Application.DisplayAlerts = True
    Next
    
    Application.DisplayAlerts = False
    Sheets("TMPREGFORM_1").Delete
    Sheets("TMPREGFORM_2").Delete
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    MsgBox "All ND-13s and their GIFs have been created and exported.", vbInformation
End Sub

Private Function NextBusinessDay(d As Date) As Date
    Dim nd As Date: nd = d + 1
    Select Case Weekday(nd, vbMonday)
        Case 6 ' Saturday
            nd = nd + 2
        Case 7 ' Sunday
            nd = nd + 1
    End Select
    NextBusinessDay = nd
End Function

Private Sub ExportSheet(ws As Worksheet, baseName As String)
    Dim fn As String
    fn = SanitizeFileName(baseName) & ".xlsx"
    Dim newWb As Workbook
    ws.Copy
    Set newWb = ActiveWorkbook
    newWb.SaveAs ThisWorkbook.Path & "\" & fn, xlOpenXMLWorkbook
    newWb.Close False
End Sub

Private Function SanitizeFileName(s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array(":", "\", "/", "?", "*", "[", "]", """")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), "")
    Next
    SanitizeFileName = Trim(s)
End Function

Private Function ParseDDMMYYYY(dateStr As String) As Date
    ' Explicitly parse DD/MM/YYYY format dates
    Dim parts() As String
    parts = Split(dateStr, "/")
    
    If UBound(parts) = 2 Then
        ' DateSerial expects (Year, Month, Day)
        ParseDDMMYYYY = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
    Else
        ' Fallback if format is unexpected
        ParseDDMMYYYY = CDate(dateStr)
    End If
End Function