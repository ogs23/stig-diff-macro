Attribute VB_Name = "ReleaseDiff"
Option Explicit

' =====================================================================
' STIG RELEASE DIFF
' =====================================================================
' What it does:
'   1. Prompts you for a CSV export that contains TWO revisions of the
'      same STIG stacked in one sheet (any STIG - column names are read
'      dynamically, nothing is hard-coded to a specific benchmark).
'   2. Imports the CSV as plain text (no auto-number/auto-date surprises),
'      saves it as a new .xlsx workbook.
'   3. Finds the "Release Info" style column (two distinct values,
'      formatted like "Release: 8 Benchmark Date: 01 Jan 2020") and a
'      unique rule-identifier column (Group ID / Vuln ID / Rule ID /
'      STIG ID - whichever exists).
'   4. Matches rules across the two revisions by that identifier and,
'      on a new "Diff" sheet, does a WORD-LEVEL diff of every column:
'        - deleted words:  red + strikethrough
'        - inserted words: green + underline
'        - any cell that changed gets a light-yellow fill so it's easy
'          to spot at a glance
'        - rules that only exist in the older revision are written in
'          solid red/strikethrough (Deleted)
'        - rules that only exist in the newer revision are written in
'          solid green/underline (New)
'   5. Writes a Notes column classifying each change (Unchanged / New /
'      Deleted / Modified - Editorial Only / Modified - Need Review),
'      naming exactly which column(s) changed.
'   6. Formats the header row (bold, light blue fill, AutoFilter) and
'      freezes the title/header rows plus every column through the
'      rule-identifier column, so it stays visible while scrolling
'      right through long prose columns.
'   7. Writes a "Summary" sheet with the same counts as the completion
'      popup (plus any duplicate-ID warnings) and a color legend,
'      persisted in the file so it travels with it if you share it.
'
' Expected input layout (this is standard for STIG CSV exports):
'   Row 1            = a classification banner (e.g. "UNCLASSIFIED"),
'                       a single populated cell - ignored automatically.
'   Row 2            = real column headers.
'   Rows 3..N-1       = rule data, two revisions stacked, distinguished
'                       by the Release Info column.
'   Row N (last row) = another classification banner - ignored
'                       automatically.
'   If your export has no banner rows at all, that's fine too - the
'   detection just falls through and uses row 1 as the header row.
'
' Run RunReleaseDiff() to start. That's the only macro you need to run.
' =====================================================================

Const SRC_SHEET_NAME As String = "Data"
Const OUT_SHEET_NAME As String = "Diff"
Const HEADER_SCAN_LIMIT As Long = 15        ' how many top rows to scan looking for the real header row
Const DIFF_TOKEN_CAP As Double = 4000000#   ' safety cap on word-diff size (old-tokens * new-tokens) before falling back to a whole-block replace

' ---------------------------------------------------------------------
' MAIN ENTRY POINT
' ---------------------------------------------------------------------
Sub RunReleaseDiff()
    Dim csvPath As Variant
    Dim wbTarget As Workbook
    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim completedOK As Boolean
    Dim oldRelease As String, newRelease As String
    Dim cntUnchanged As Long, cntEditorial As Long, cntReview As Long, cntNew As Long, cntDeleted As Long
    Dim cntDupOld As Long, cntDupNew As Long
    Dim sourceLabel As String

    completedOK = False

    csvPath = Application.GetOpenFilename( _
        "CSV Files (*.csv),*.csv", , _
        "Select the CSV export containing BOTH STIG revisions (Cancel to reuse the active workbook's '" & SRC_SHEET_NAME & "' sheet)")

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    If csvPath = False Then
        ' No file chosen - fall back to the active workbook if it already has a Data sheet
        On Error Resume Next
        Set wsSrc = ActiveWorkbook.Worksheets(SRC_SHEET_NAME)
        On Error GoTo ErrHandler
        If wsSrc Is Nothing Then
            MsgBox "No CSV selected, and the active workbook has no '" & SRC_SHEET_NAME & "' sheet to fall back on. Nothing to do.", vbExclamation
            GoTo CleanExit
        End If
        Set wbTarget = ActiveWorkbook
        sourceLabel = "(existing workbook) " & wbTarget.Name
    Else
        Set wbTarget = Workbooks.Open(Filename:=CStr(csvPath), Local:=True)
        Set wsSrc = wbTarget.Worksheets(1)
        wsSrc.Name = SRC_SHEET_NAME
        RD_NeutralizeAutoTypes wsSrc
        sourceLabel = CStr(csvPath)

        Dim xlsxPath As String
        xlsxPath = RD_SwapExtension(CStr(csvPath), ".xlsx")
        wbTarget.SaveAs Filename:=xlsxPath, FileFormat:=xlOpenXMLWorkbook
    End If

    Set wsOut = RD_GetOrCreateSheet(wbTarget, OUT_SHEET_NAME)
    wsOut.Cells.Clear

    ' ---- locate structure (skips classification banner rows automatically) ----
    Dim headerRow As Long, firstDataRow As Long, lastDataRow As Long
    headerRow = RD_FindHeaderRow(wsSrc)
    firstDataRow = headerRow + 1
    lastDataRow = RD_FindLastDataRow(wsSrc, firstDataRow)

    If lastDataRow < firstDataRow Then
        MsgBox "No data rows found on '" & wsSrc.Name & "' below the header row (row " & headerRow & ").", vbCritical
        GoTo CleanExit
    End If

    Dim lastCol As Long
    lastCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column

    Dim releaseCol As Long, keyCol As Long
    releaseCol = RD_FindReleaseColumn(wsSrc, headerRow, firstDataRow, lastCol)
    keyCol = RD_FindKeyColumn(wsSrc, headerRow, lastCol)

    If releaseCol = 0 Or keyCol = 0 Then
        Dim diagMsg As String
        diagMsg = "Header row detected: row " & headerRow & vbCrLf & _
                  "Headers found there: " & RD_JoinRowValues(wsSrc, headerRow, lastCol) & vbCrLf & vbCrLf

        If releaseCol = 0 Then
            diagMsg = diagMsg & "Could not find a 'Release Info' style column (expected text like 'Release: 8 Benchmark Date: 01 Jan 2020' somewhere in a column). Add/rename that column and re-run."
        Else
            diagMsg = diagMsg & "Could not find a unique rule-identifier column (looked for headers named Group ID / Vuln ID / Rule ID / STIG ID / Legacy ID). Rename your ID column to one of these and re-run."
        End If

        MsgBox diagMsg, vbCritical
        GoTo CleanExit
    End If

    ' ---- build the list of output columns: every source column except Release Info ----
    Dim srcCols() As Long, srcHeaders() As String
    Dim n As Long, c As Long
    ReDim srcCols(0 To lastCol - 1)
    ReDim srcHeaders(0 To lastCol - 1)
    n = 0
    For c = 1 To lastCol
        If c <> releaseCol Then
            srcCols(n) = c
            srcHeaders(n) = Trim(RD_SafeStr(wsSrc.Cells(headerRow, c).Value))
            n = n + 1
        End If
    Next c
    ReDim Preserve srcCols(0 To n - 1)
    ReDim Preserve srcHeaders(0 To n - 1)

    Dim notesHeaderText As String
    notesHeaderText = "Notes"
    For c = 0 To n - 1
        If LCase(srcHeaders(c)) = "notes" Then notesHeaderText = "Diff Notes": Exit For
    Next c

    ' ---- discover the two distinct releases and which is older ----
    Dim releaseKeys As New Collection, releaseDates As New Collection
    Dim r As Long, relTxt As String, relDate As Date

    For r = firstDataRow To lastDataRow
        relTxt = Trim(RD_SafeStr(wsSrc.Cells(r, releaseCol).Value))
        If relTxt <> "" Then
            If Not RD_KeyExists(releaseKeys, relTxt) Then
                relDate = RD_ParseBenchmarkDate(relTxt)
                releaseKeys.Add relTxt, relTxt
                releaseDates.Add relDate, relTxt
            End If
        End If
    Next r

    If releaseKeys.Count <> 2 Then
        Dim msg As String, rk As Variant
        msg = "Expected exactly 2 distinct '" & wsSrc.Cells(headerRow, releaseCol).Value & "' values, found " & releaseKeys.Count & ":" & vbCrLf
        For Each rk In releaseKeys
            msg = msg & "  - " & rk & vbCrLf
        Next rk
        MsgBox msg, vbCritical
        GoTo CleanExit
    End If

    Dim relArr(1 To 2) As String, dateArr(1 To 2) As Date, idx As Integer
    idx = 1
    For Each rk In releaseKeys
        relArr(idx) = CStr(rk)
        dateArr(idx) = releaseDates(CStr(rk))
        idx = idx + 1
    Next rk

    If dateArr(1) <= dateArr(2) Then
        oldRelease = relArr(1): newRelease = relArr(2)
    Else
        oldRelease = relArr(2): newRelease = relArr(1)
    End If

    ' ---- index rows by revision + key ----
    Dim oldRows As New Collection, newRows As New Collection, gTxt As String
    For r = firstDataRow To lastDataRow
        relTxt = Trim(RD_SafeStr(wsSrc.Cells(r, releaseCol).Value))
        gTxt = Trim(RD_SafeStr(wsSrc.Cells(r, keyCol).Value))
        If gTxt <> "" Then
            If relTxt = oldRelease Then
                If RD_KeyExists(oldRows, gTxt) Then
                    cntDupOld = cntDupOld + 1
                Else
                    oldRows.Add r, gTxt
                End If
            ElseIf relTxt = newRelease Then
                If RD_KeyExists(newRows, gTxt) Then
                    cntDupNew = cntDupNew + 1
                Else
                    newRows.Add r, gTxt
                End If
            End If
        End If
    Next r

    ' ---- title banner + headers ----
    Dim outColCount As Long: outColCount = n + 1 ' + Notes
    Dim notesCol As Long: notesCol = outColCount

    Dim keyOutCol As Long: keyOutCol = 1
    For i = 0 To n - 1
        If srcCols(i) = keyCol Then keyOutCol = i + 1: Exit For
    Next i

    ' Force every output cell to Text format up front so nothing (IDs,
    ' version-like strings, dates embedded in prose) gets silently
    ' reinterpreted as a number/date when we write it.
    wsOut.Range(wsOut.Columns(1), wsOut.Columns(outColCount)).NumberFormat = "@"

    wsOut.Cells(1, 1).Value = "STIG Comparison   |   Older: " & oldRelease & "   |   Newer: " & newRelease
    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, outColCount))
        .Merge
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    wsOut.Rows(1).RowHeight = 22

    Dim i As Long
    For i = 0 To n - 1
        wsOut.Cells(2, i + 1).Value = srcHeaders(i)
    Next i
    wsOut.Cells(2, notesCol).Value = notesHeaderText

    ' ---- write data rows ----
    Dim criticalKeywords() As String
    criticalKeywords = Split("severity,stig id,rule id,cci,weight,vuln id", ",")

    Dim outRow As Long: outRow = 3
    Dim totalRows As Long: totalRows = lastDataRow - firstDataRow + 1

    For r = firstDataRow To lastDataRow
        If (r - firstDataRow) Mod 25 = 0 Then
            Application.StatusBar = "Comparing rules... " & (r - firstDataRow + 1) & " of " & totalRows
        End If

        relTxt = Trim(RD_SafeStr(wsSrc.Cells(r, releaseCol).Value))
        gTxt = Trim(RD_SafeStr(wsSrc.Cells(r, keyCol).Value))

        If relTxt = oldRelease And gTxt <> "" Then
            If Not RD_KeyExists(newRows, gTxt) Then
                RD_ProcessWholeRow wsSrc, wsOut, outRow, r, srcCols, notesCol, "del", "Deleted - not present in " & newRelease
                cntDeleted = cntDeleted + 1
            Else
                RD_ProcessMatchedRow wsSrc, wsOut, outRow, r, CLng(newRows(gTxt)), srcCols, srcHeaders, keyCol, notesCol, _
                    criticalKeywords, cntUnchanged, cntEditorial, cntReview
            End If
            outRow = outRow + 1
        ElseIf relTxt = newRelease And gTxt <> "" Then
            If Not RD_KeyExists(oldRows, gTxt) Then
                RD_ProcessWholeRow wsSrc, wsOut, outRow, r, srcCols, notesCol, "ins", "New - not present in " & oldRelease
                cntNew = cntNew + 1
                outRow = outRow + 1
            End If
        End If
    Next r

    Application.StatusBar = "Formatting output..."
    RD_FormatOutput wsOut, outRow - 1, outColCount, keyOutCol

    RD_WriteSummarySheet wbTarget, sourceLabel, oldRelease, newRelease, _
        cntUnchanged, cntEditorial, cntReview, cntNew, cntDeleted, cntDupOld, cntDupNew

    wbTarget.Save
    completedOK = True

CleanExit:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    If completedOK Then
        Dim summaryMsg As String
        summaryMsg = "Release diff complete." & vbCrLf & _
               "Older: " & oldRelease & vbCrLf & "Newer: " & newRelease & vbCrLf & vbCrLf & _
               "Unchanged: " & cntUnchanged & vbCrLf & _
               "Modified (editorial only): " & cntEditorial & vbCrLf & _
               "Modified (needs review): " & cntReview & vbCrLf & _
               "New rules: " & cntNew & vbCrLf & _
               "Deleted rules: " & cntDeleted
        If cntDupOld > 0 Or cntDupNew > 0 Then
            summaryMsg = summaryMsg & vbCrLf & vbCrLf & _
                "Note: duplicate rule IDs were found (" & cntDupOld & " in the older release, " & _
                cntDupNew & " in the newer) - only the first occurrence of each was used. See the Summary sheet."
        End If
        summaryMsg = summaryMsg & vbCrLf & vbCrLf & "Full breakdown saved on the Summary sheet."
        MsgBox summaryMsg, vbInformation
    End If
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ---------------------------------------------------------------------
' CSV -> XLSX IMPORT
' ---------------------------------------------------------------------
' The CSV itself is opened with Workbooks.Open, i.e. Excel's own native
' CSV parser - not a hand-configured Text Import (QueryTables), which
' has a long-standing bug where embedded newlines inside quoted fields
' (very common in STIG Discussion/Check Content text) get treated as
' real row breaks, shredding a single rule's row across several rows
' and throwing off everything below it (including which row is really
' the header row). Workbooks.Open handles quoted embedded commas and
' newlines correctly.
'
' The one thing Workbooks.Open can still do is auto-type a cell that
' looks like a number or date (e.g. a bare version number) instead of
' leaving it as text. RD_NeutralizeAutoTypes below fixes that up
' afterward without touching anything that was already text.
Sub RD_NeutralizeAutoTypes(ws As Worksheet)
    Dim usedRng As Range
    Set usedRng = ws.UsedRange
    If usedRng.Cells.Count = 0 Then Exit Sub

    Dim rowStart As Long, colStart As Long, lastR As Long, lastC As Long
    rowStart = usedRng.Row
    colStart = usedRng.Column
    lastR = rowStart + usedRng.Rows.Count - 1
    lastC = colStart + usedRng.Columns.Count - 1

    Dim dataArr As Variant
    dataArr = ws.Range(ws.Cells(rowStart, colStart), ws.Cells(lastR, lastC)).Value2

    Dim r As Long, c As Long, v As Variant, vt As VbVarType
    For r = 1 To UBound(dataArr, 1)
        For c = 1 To UBound(dataArr, 2)
            v = dataArr(r, c)
            vt = VarType(v)
            If vt = vbDouble Or vt = vbSingle Or vt = vbInteger Or vt = vbLong Or vt = vbDate Or vt = vbCurrency Then
                ' Was auto-typed as a number/date - grab the exact displayed text instead.
                dataArr(r, c) = ws.Cells(rowStart + r - 1, colStart + c - 1).Text
            End If
        Next c
    Next r

    With ws.Range(ws.Cells(rowStart, colStart), ws.Cells(lastR, lastC))
        .NumberFormat = "@"
        .Value = dataArr
    End With
End Sub

Function RD_SwapExtension(path As String, newExt As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(path, ".")
    If dotPos = 0 Then
        RD_SwapExtension = path & newExt
    Else
        RD_SwapExtension = Left(path, dotPos - 1) & newExt
    End If
End Function

' ---------------------------------------------------------------------
' STRUCTURE DETECTION
' ---------------------------------------------------------------------
Function RD_GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set RD_GetOrCreateSheet = ws
End Function

' Finds the real header row. Prefers a row that both (a) has several
' populated cells (rules out a 1-cell classification banner) and (b)
' actually looks like STIG column headers (rules out a banner or other
' junk row that happens to have a few cells populated). Falls back to
' the old "just count cells" heuristic if nothing scores as a header,
' so unusual exports still get a best-effort guess instead of a hard
' failure.
Function RD_FindHeaderRow(ws As Worksheet) As Long
    Dim keywords() As String
    keywords = Split("group id,severity,release,rule id,stig id,benchmark,check content,fix text,discussion,vuln,rule title,cci,version,weight", ",")

    Dim r As Long, fallbackRow As Long
    fallbackRow = 0

    For r = 1 To HEADER_SCAN_LIMIT
        If Application.WorksheetFunction.CountA(ws.Rows(r)) >= 3 Then
            If fallbackRow = 0 Then fallbackRow = r
            If RD_ScoreHeaderRow(ws, r, keywords) >= 2 Then
                RD_FindHeaderRow = r
                Exit Function
            End If
        End If
    Next r

    If fallbackRow > 0 Then
        RD_FindHeaderRow = fallbackRow
    Else
        RD_FindHeaderRow = 1 ' last resort: assume no banner row is present
    End If
End Function

' Counts how many recognizable STIG header keywords appear among row r's
' populated cells.
Function RD_ScoreHeaderRow(ws As Worksheet, r As Long, keywords() As String) As Long
    Dim lastC As Long, c As Long, cellTxt As String, score As Long, kw As Variant
    lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
    If lastC > 60 Then lastC = 60 ' cap the scan width, headers never run this wide

    score = 0
    For c = 1 To lastC
        cellTxt = LCase(Trim(RD_SafeStr(ws.Cells(r, c).Value)))
        If cellTxt <> "" Then
            For Each kw In keywords
                If InStr(cellTxt, CStr(kw)) > 0 Then
                    score = score + 1
                    Exit For
                End If
            Next kw
        End If
    Next c
    RD_ScoreHeaderRow = score
End Function

' Joins a row's populated header values for diagnostic error messages.
Function RD_JoinRowValues(ws As Worksheet, r As Long, lastCol As Long) As String
    Dim c As Long, result As String
    For c = 1 To lastCol
        result = result & "[" & RD_SafeStr(ws.Cells(r, c).Value) & "] "
    Next c
    RD_JoinRowValues = Trim(result)
End Function

' Finds the last real data row, trimming off a trailing classification-
' banner row (again, a row with 0-1 populated cells).
Function RD_FindLastDataRow(ws As Worksheet, firstDataRow As Long) As Long
    Dim lastRow As Long
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row
    Do While lastRow > firstDataRow And Application.WorksheetFunction.CountA(ws.Rows(lastRow)) <= 1
        lastRow = lastRow - 1
    Loop
    If lastRow < firstDataRow Then
        RD_FindLastDataRow = firstDataRow - 1 ' signals "no data"
    Else
        RD_FindLastDataRow = lastRow
    End If
End Function

' Finds the Release Info column: first by header name, then (fallback)
' by scanning early data rows for the "Release: ... Benchmark Date: ..."
' pattern, so this still works if the header is named/spelled differently.
Function RD_FindReleaseColumn(ws As Worksheet, headerRow As Long, firstDataRow As Long, lastCol As Long) As Long
    Dim c As Long, hdr As String
    For c = 1 To lastCol
        hdr = LCase(Trim(RD_SafeStr(ws.Cells(headerRow, c).Value)))
        If InStr(hdr, "release") > 0 Then
            RD_FindReleaseColumn = c
            Exit Function
        End If
    Next c

    Dim r As Long, sampleEnd As Long, v As String
    sampleEnd = WorksheetFunction.Min(firstDataRow + 10, ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row)
    For c = 1 To lastCol
        For r = firstDataRow To sampleEnd
            v = RD_SafeStr(ws.Cells(r, c).Value)
            If InStr(1, v, "Release:", vbTextCompare) > 0 And InStr(1, v, "Benchmark Date:", vbTextCompare) > 0 Then
                RD_FindReleaseColumn = c
                Exit Function
            End If
        Next r
    Next c
    RD_FindReleaseColumn = 0
End Function

' Finds a stable per-rule identifier column, trying the common STIG
' export header names in priority order.
Function RD_FindKeyColumn(ws As Worksheet, headerRow As Long, lastCol As Long) As Long
    Dim priorities() As String
    priorities = Split("Group ID,Vuln ID,Rule ID,STIG ID,Legacy ID", ",")
    Dim p As Variant, c As Long
    For Each p In priorities
        For c = 1 To lastCol
            If LCase(Trim(RD_SafeStr(ws.Cells(headerRow, c).Value))) = LCase(CStr(p)) Then
                RD_FindKeyColumn = c
                Exit Function
            End If
        Next c
    Next p
    RD_FindKeyColumn = 0
End Function

' ---------------------------------------------------------------------
' SMALL UTILITIES
' ---------------------------------------------------------------------
Function RD_KeyExists(col As Collection, key As String) As Boolean
    Dim x As Variant
    On Error GoTo NotFound
    x = col(key)
    RD_KeyExists = True
    Exit Function
NotFound:
    RD_KeyExists = False
End Function

Function RD_SafeStr(v As Variant) As String
    If IsError(v) Then
        RD_SafeStr = ""
    ElseIf IsNull(v) Then
        RD_SafeStr = ""
    Else
        RD_SafeStr = CStr(v)
    End If
End Function

Function RD_ParseBenchmarkDate(releaseText As String) As Date
    Dim p As Long
    Dim datePart As String
    p = InStr(1, releaseText, "Benchmark Date:", vbTextCompare)
    If p = 0 Then
        Err.Raise vbObjectError + 1, "RD_ParseBenchmarkDate", _
            "No 'Benchmark Date:' found in Release Info value: " & releaseText
    End If
    datePart = Trim(Mid(releaseText, p + Len("Benchmark Date:")))

    On Error GoTo ParseFail
    RD_ParseBenchmarkDate = DateValue(datePart)
    Exit Function

ParseFail:
    Err.Raise vbObjectError + 2, "RD_ParseBenchmarkDate", _
        "Could not parse date '" & datePart & "' from Release Info: " & releaseText
End Function

Function RD_IsCriticalHeader(headerName As String, keywords() As String) As Boolean
    Dim k As Variant, h As String
    h = LCase(headerName)
    For Each k In keywords
        If InStr(h, LCase(CStr(k))) > 0 Then
            RD_IsCriticalHeader = True
            Exit Function
        End If
    Next k
    RD_IsCriticalHeader = False
End Function

Function RD_ClassifyRow(changedHeaders As Collection, criticalChanged As Boolean, allNormalizedEqual As Boolean) As String
    If changedHeaders.Count = 0 Then
        RD_ClassifyRow = "Unchanged"
        Exit Function
    End If

    Dim lst As String, h As Variant
    For Each h In changedHeaders
        lst = RD_AppendField(lst, CStr(h))
    Next h

    If criticalChanged Or Not allNormalizedEqual Then
        RD_ClassifyRow = "Modified - Need Review (" & lst & ")"
    Else
        RD_ClassifyRow = "Modified - Editorial Only (" & lst & ")"
    End If
End Function

Function RD_AppendField(existingList As String, fieldName As String) As String
    If existingList = "" Then
        RD_AppendField = fieldName
    Else
        RD_AppendField = existingList & ", " & fieldName
    End If
End Function

' Normalizes whitespace/punctuation/case so purely cosmetic edits
' (re-wrapped text, added periods, re-cased words) don't get flagged
' as needing review.
Function RD_Normalize(s As String) As String
    Dim t As String
    Dim punct As Variant, p As Variant

    t = LCase(s)
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")

    punct = Array(".", ",", ";", ":", """", "'", "(", ")", "-")
    For Each p In punct
        t = Replace(t, p, "")
    Next p

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    RD_Normalize = Trim(t)
End Function

' ---------------------------------------------------------------------
' ROW PROCESSORS
' ---------------------------------------------------------------------

' A rule that exists in both revisions: diff every column (except the
' key column) word-by-word, classify, and write the Notes entry.
Sub RD_ProcessMatchedRow(wsSrc As Worksheet, wsOut As Worksheet, outRow As Long, oldRow As Long, newRow As Long, _
                          srcCols() As Long, srcHeaders() As String, keyColIdx As Long, notesCol As Long, _
                          criticalKeywords() As String, _
                          ByRef cntUnchanged As Long, ByRef cntEditorial As Long, ByRef cntReview As Long)
    Dim i As Long, outCol As Long
    Dim oldVal As String, newVal As String
    Dim outCell As Range
    Dim changedHeaders As New Collection
    Dim criticalChanged As Boolean: criticalChanged = False
    Dim allNormEqual As Boolean: allNormEqual = True

    For i = LBound(srcCols) To UBound(srcCols)
        outCol = i - LBound(srcCols) + 1
        oldVal = RD_SafeStr(wsSrc.Cells(oldRow, srcCols(i)).Value)
        newVal = RD_SafeStr(wsSrc.Cells(newRow, srcCols(i)).Value)
        Set outCell = wsOut.Cells(outRow, outCol)

        If srcCols(i) = keyColIdx Then
            outCell.Value = oldVal
        ElseIf oldVal = newVal Then
            outCell.Value = oldVal
        Else
            RD_WriteWordDiff outCell, oldVal, newVal
            changedHeaders.Add srcHeaders(i), srcHeaders(i)
            If RD_IsCriticalHeader(srcHeaders(i), criticalKeywords) Then criticalChanged = True
            If RD_Normalize(oldVal) <> RD_Normalize(newVal) Then allNormEqual = False
        End If
    Next i

    Dim noteTxt As String
    noteTxt = RD_ClassifyRow(changedHeaders, criticalChanged, allNormEqual)
    wsOut.Cells(outRow, notesCol).Value = noteTxt
    RD_ColorNote wsOut.Cells(outRow, notesCol), noteTxt

    Select Case True
        Case noteTxt = "Unchanged": cntUnchanged = cntUnchanged + 1
        Case InStr(noteTxt, "Editorial Only") > 0: cntEditorial = cntEditorial + 1
        Case Else: cntReview = cntReview + 1
    End Select
End Sub

' A rule that only exists in one revision (New or Deleted): render the
' whole row in solid redline color, no cell-by-cell diff needed.
Sub RD_ProcessWholeRow(wsSrc As Worksheet, wsOut As Worksheet, outRow As Long, srcRow As Long, _
                        srcCols() As Long, notesCol As Long, colorType As String, noteTxt As String)
    Dim i As Long, outCol As Long, v As String

    For i = LBound(srcCols) To UBound(srcCols)
        outCol = i - LBound(srcCols) + 1
        v = RD_SafeStr(wsSrc.Cells(srcRow, srcCols(i)).Value)
        With wsOut.Cells(outRow, outCol)
            .Value = v
            If Len(v) > 0 Then
                If colorType = "del" Then
                    .Font.Color = RGB(192, 0, 0)
                    .Font.Strikethrough = True
                ElseIf colorType = "ins" Then
                    .Font.Color = RGB(0, 128, 0)
                    .Font.Underline = xlUnderlineStyleSingle
                End If
            End If
        End With
    Next i

    wsOut.Cells(outRow, notesCol).Value = noteTxt
    RD_ColorNote wsOut.Cells(outRow, notesCol), noteTxt

    Dim tintColor As Long
    If colorType = "del" Then tintColor = RGB(255, 235, 235) Else tintColor = RGB(235, 255, 235)
    wsOut.Range(wsOut.Cells(outRow, 1), wsOut.Cells(outRow, notesCol - 1)).Interior.Color = tintColor
End Sub

Sub RD_ColorNote(targetCell As Range, noteTxt As String)
    If InStr(1, noteTxt, "Deleted", vbTextCompare) = 1 Then
        targetCell.Interior.Color = RGB(255, 199, 206)
    ElseIf InStr(1, noteTxt, "New", vbTextCompare) = 1 Then
        targetCell.Interior.Color = RGB(198, 239, 206)
    ElseIf InStr(1, noteTxt, "Modified - Need Review", vbTextCompare) = 1 Then
        targetCell.Interior.Color = RGB(255, 235, 156)
    ElseIf InStr(1, noteTxt, "Modified - Editorial Only", vbTextCompare) = 1 Then
        targetCell.Interior.Color = RGB(221, 235, 247)
    End If
End Sub

' ---------------------------------------------------------------------
' WORD-LEVEL DIFF ENGINE
' ---------------------------------------------------------------------

' Writes a merged, redlined version of oldText -> newText into `cell`:
' unchanged words in normal text, deleted words in red+strikethrough,
' inserted words in green+underline. Also flags the cell with a light
' yellow fill so a changed cell is obvious even before reading it.
Sub RD_WriteWordDiff(cell As Range, oldText As String, newText As String)
    Dim oldArr() As String, newArr() As String
    oldArr = RD_Tokenize(oldText)
    newArr = RD_Tokenize(newText)

    Dim nOld As Long, nNew As Long
    nOld = UBound(oldArr) + 1
    nNew = UBound(newArr) + 1

    Dim fullText As String, curType As String, curStart As Long, curLen As Long
    Dim runStarts() As Long, runLens() As Long, runTypes() As String, runCount As Long
    curType = "": curStart = 1: curLen = 0: fullText = "": runCount = 0
    ReDim runStarts(0 To 0): ReDim runLens(0 To 0): ReDim runTypes(0 To 0)

    If CDbl(nOld) * CDbl(nNew) > DIFF_TOKEN_CAP Then
        ' Too large for a token-level diff - fall back to a whole-block replace.
        If nOld > 0 Then RD_PushToken fullText, curType, curStart, curLen, runStarts, runLens, runTypes, runCount, "del", oldText
        If nNew > 0 Then RD_PushToken fullText, curType, curStart, curLen, runStarts, runLens, runTypes, runCount, "ins", newText
    Else
        Dim dp() As Long
        ReDim dp(0 To nOld, 0 To nNew)
        Dim i As Long, j As Long
        For i = nOld - 1 To 0 Step -1
            For j = nNew - 1 To 0 Step -1
                If oldArr(i) = newArr(j) Then
                    dp(i, j) = dp(i + 1, j + 1) + 1
                ElseIf dp(i + 1, j) >= dp(i, j + 1) Then
                    dp(i, j) = dp(i + 1, j)
                Else
                    dp(i, j) = dp(i, j + 1)
                End If
            Next j
        Next i

        i = 0: j = 0
        Do While i < nOld And j < nNew
            If oldArr(i) = newArr(j) Then
                RD_PushToken fullText, curType, curStart, curLen, runStarts, runLens, runTypes, runCount, "eq", oldArr(i)
                i = i + 1: j = j + 1
            ElseIf dp(i + 1, j) >= dp(i, j + 1) Then
                RD_PushToken fullText, curType, curStart, curLen, runStarts, runLens, runTypes, runCount, "del", oldArr(i)
                i = i + 1
            Else
                RD_PushToken fullText, curType, curStart, curLen, runStarts, runLens, runTypes, runCount, "ins", newArr(j)
                j = j + 1
            End If
        Loop
        Do While i < nOld
            RD_PushToken fullText, curType, curStart, curLen, runStarts, runLens, runTypes, runCount, "del", oldArr(i)
            i = i + 1
        Loop
        Do While j < nNew
            RD_PushToken fullText, curType, curStart, curLen, runStarts, runLens, runTypes, runCount, "ins", newArr(j)
            j = j + 1
        Loop
    End If

    RD_FlushRun runStarts, runLens, runTypes, runCount, curType, curStart, curLen

    cell.Value = fullText
    cell.Interior.Color = RGB(255, 242, 204) ' flag: this cell changed

    Dim k As Long
    For k = 0 To runCount - 1
        With cell.Characters(Start:=runStarts(k), Length:=runLens(k)).Font
            Select Case runTypes(k)
                Case "del"
                    .Color = RGB(192, 0, 0)
                    .Strikethrough = True
                    .Underline = xlUnderlineStyleNone
                Case "ins"
                    .Color = RGB(0, 128, 0)
                    .Strikethrough = False
                    .Underline = xlUnderlineStyleSingle
                Case Else ' "eq"
                    .ColorIndex = xlAutomatic
                    .Strikethrough = False
                    .Underline = xlUnderlineStyleNone
            End Select
        End With
    Next k
End Sub

' Splits text into alternating runs of non-whitespace / whitespace,
' preserving every character so the pieces reconstruct exactly - this
' is what lets the diff show real inserted/deleted words, not just
' "this whole field is different".
Function RD_Tokenize(ByVal s As String) As String()
    Dim n As Long: n = Len(s)
    Dim tokens() As String

    If n = 0 Then
        ReDim tokens(0 To -1)
        RD_Tokenize = tokens
        Exit Function
    End If

    ReDim tokens(0 To n - 1)
    Dim cnt As Long: cnt = 0
    Dim i As Long: i = 1
    Dim ch As String, isSpace As Boolean, startPos As Long

    Do While i <= n
        ch = Mid$(s, i, 1)
        isSpace = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf)
        startPos = i
        i = i + 1
        Do While i <= n
            ch = Mid$(s, i, 1)
            If ((ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf) <> isSpace) Then Exit Do
            i = i + 1
        Loop
        tokens(cnt) = Mid$(s, startPos, i - startPos)
        cnt = cnt + 1
    Loop

    ReDim Preserve tokens(0 To cnt - 1)
    RD_Tokenize = tokens
End Function

' Appends one token to the run currently being built, starting a new
' run (and flushing the previous one) whenever the op type changes.
Sub RD_PushToken(ByRef fullText As String, ByRef curType As String, ByRef curStart As Long, ByRef curLen As Long, _
                  ByRef runStarts() As Long, ByRef runLens() As Long, ByRef runTypes() As String, ByRef runCount As Long, _
                  ByVal opType As String, ByVal tokText As String)
    If opType <> curType Then
        RD_FlushRun runStarts, runLens, runTypes, runCount, curType, curStart, curLen
        curType = opType
        curStart = Len(fullText) + 1
        curLen = 0
    End If
    curLen = curLen + Len(tokText)
    fullText = fullText & tokText
End Sub

Sub RD_FlushRun(ByRef runStarts() As Long, ByRef runLens() As Long, ByRef runTypes() As String, ByRef runCount As Long, _
                ByVal curType As String, ByVal curStart As Long, ByVal curLen As Long)
    If curType = "" Or curLen = 0 Then Exit Sub
    ReDim Preserve runStarts(0 To runCount)
    ReDim Preserve runLens(0 To runCount)
    ReDim Preserve runTypes(0 To runCount)
    runStarts(runCount) = curStart
    runLens(runCount) = curLen
    runTypes(runCount) = curType
    runCount = runCount + 1
End Sub

' ---------------------------------------------------------------------
' OUTPUT FORMATTING
' ---------------------------------------------------------------------
Sub RD_FormatOutput(ws As Worksheet, lastDataRow As Long, outColCount As Long, keyOutCol As Long)
    Const headerRowNum As Long = 2

    Dim headerRange As Range
    Set headerRange = ws.Range(ws.Cells(headerRowNum, 1), ws.Cells(headerRowNum, outColCount))
    With headerRange
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230) ' light blue
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(headerRowNum).RowHeight = 20

    Dim c As Long
    For c = 1 To outColCount
        ws.Columns(c).ColumnWidth = 32
    Next c
    ws.Columns(outColCount).ColumnWidth = 40 ' Notes - a bit wider

    If lastDataRow >= headerRowNum + 1 Then
        With ws.Range(ws.Cells(headerRowNum + 1, 1), ws.Cells(lastDataRow, outColCount))
            .WrapText = True
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlLeft
        End With
        ws.Rows(CStr(headerRowNum + 1) & ":" & CStr(lastDataRow)).AutoFit
    End If

    Dim bottomRow As Long
    bottomRow = IIf(lastDataRow >= headerRowNum, lastDataRow, headerRowNum)
    With ws.Range(ws.Cells(headerRowNum, 1), ws.Cells(bottomRow, outColCount)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With

    headerRange.AutoFilter

    ws.Activate
    ' Freeze the title/header rows AND every column up through the rule
    ' identifier, so it stays in view while scrolling right through the
    ' long prose columns (Fix Text, Discussion, etc.).
    ws.Cells(headerRowNum + 1, keyOutCol + 1).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
End Sub

' ---------------------------------------------------------------------
' SUMMARY SHEET
' ---------------------------------------------------------------------
' Persists the run's counts and a color legend into the workbook itself
' (rather than just a popup that disappears), since this file is meant
' to be handed off/shared as a compliance artifact.
Sub RD_WriteSummarySheet(wb As Workbook, sourceLabel As String, oldRelease As String, newRelease As String, _
                          cntUnchanged As Long, cntEditorial As Long, cntReview As Long, cntNew As Long, cntDeleted As Long, _
                          cntDupOld As Long, cntDupNew As Long)
    Dim ws As Worksheet
    Set ws = RD_GetOrCreateSheet(wb, "Summary")
    ws.Cells.Clear

    With ws.Range("A1:B1")
        .Merge
        .Value = "STIG Release Comparison - Summary"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22

    ws.Range("A3").Value = "Source:"
    ws.Range("B3").Value = sourceLabel
    ws.Range("A4").Value = "Older release:"
    ws.Range("B4").Value = oldRelease
    ws.Range("A5").Value = "Newer release:"
    ws.Range("B5").Value = newRelease
    ws.Range("A6").Value = "Generated:"
    ws.Range("B6").Value = Format(Now, "yyyy-mm-dd hh:mm")
    ws.Range("A3:A6").Font.Bold = True

    Dim r As Long
    r = 8
    ws.Cells(r, 1).Value = "Change Summary"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
        .Font.Bold = True
        .Interior.Color = RGB(173, 216, 230)
    End With
    ws.Cells(r, 1).Value = "Category"
    ws.Cells(r, 2).Value = "Count"
    r = r + 1

    RD_WriteSummaryRow ws, r, "Unchanged", cntUnchanged: r = r + 1
    RD_WriteSummaryRow ws, r, "Modified - Editorial Only", cntEditorial: r = r + 1
    RD_WriteSummaryRow ws, r, "Modified - Need Review", cntReview: r = r + 1
    RD_WriteSummaryRow ws, r, "New rules", cntNew: r = r + 1
    RD_WriteSummaryRow ws, r, "Deleted rules", cntDeleted: r = r + 1
    RD_WriteSummaryRow ws, r, "Total rules compared", cntUnchanged + cntEditorial + cntReview + cntNew + cntDeleted
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Bold = True
    r = r + 1

    If cntDupOld > 0 Or cntDupNew > 0 Then
        r = r + 1
        RD_WriteSummaryRow ws, r, "Duplicate IDs skipped - older release (first occurrence kept)", cntDupOld: r = r + 1
        RD_WriteSummaryRow ws, r, "Duplicate IDs skipped - newer release (first occurrence kept)", cntDupNew: r = r + 1
    End If

    r = r + 2
    ws.Cells(r, 1).Value = "Color Legend"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    With ws.Cells(r, 1)
        .Value = "(cell fill)"
        .Interior.Color = RGB(255, 242, 204)
    End With
    ws.Cells(r, 2).Value = "Cell changed between releases - the redlined wording is inside it"
    r = r + 1

    With ws.Cells(r, 1)
        .Value = "Deleted wording"
        .Font.Color = RGB(192, 0, 0)
        .Font.Strikethrough = True
    End With
    ws.Cells(r, 2).Value = "Word removed going from the older to the newer release"
    r = r + 1

    With ws.Cells(r, 1)
        .Value = "Inserted wording"
        .Font.Color = RGB(0, 128, 0)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    ws.Cells(r, 2).Value = "Word added going from the older to the newer release"
    r = r + 1

    RD_WriteLegendRow ws, r, RGB(255, 199, 206), "Deleted rule (Notes cell)", "Rule exists only in the older release": r = r + 1
    RD_WriteLegendRow ws, r, RGB(198, 239, 206), "New rule (Notes cell)", "Rule exists only in the newer release": r = r + 1
    RD_WriteLegendRow ws, r, RGB(255, 235, 156), "Modified - Need Review (Notes cell)", "Substantive change - worth a human look": r = r + 1
    RD_WriteLegendRow ws, r, RGB(221, 235, 247), "Modified - Editorial Only (Notes cell)", "Wording/formatting change only, no real meaning change": r = r + 1

    ws.Columns("A").ColumnWidth = 42
    ws.Columns("B").ColumnWidth = 50
    ws.Range(ws.Cells(3, 1), ws.Cells(r, 2)).WrapText = True

    ws.Activate
    ws.Cells(1, 1).Select
End Sub

Sub RD_WriteSummaryRow(ws As Worksheet, r As Long, label As String, cnt As Long)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Value = cnt
End Sub

Sub RD_WriteLegendRow(ws As Worksheet, r As Long, fillColor As Long, sampleText As String, meaning As String)
    With ws.Cells(r, 1)
        .Value = sampleText
        .Interior.Color = fillColor
    End With
    ws.Cells(r, 2).Value = meaning
End Sub
