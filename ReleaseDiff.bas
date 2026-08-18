Sub RunReleaseDiff()
    ' Compares two revisions of the same STIG list that live in a single
    ' sheet, distinguished by a "Release Info" column such as:
    '   "Release: 8 Benchmark Date: 01 Jan 2020"
    ' Rows are matched across revisions by Group ID. The result is a
    ' change category written into the Notes column of a copy on a
    ' separate output sheet - the source sheet is never modified.
    '
    ' Update these two constants if your sheet names differ.
    Const SRC_SHEET_NAME As String = "Data"
    Const OUT_SHEET_NAME As String = "Diff"

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim oldRelease As String
    Dim newRelease As String
    oldRelease = ""
    newRelease = ""

    Set wsSrc = Sheets(SRC_SHEET_NAME)
    Set wsOut = RD_GetOrCreateSheet(OUT_SHEET_NAME)
    wsOut.Cells.Clear

    Dim headers() As String
    headers = Split("Release Info,Group ID,Severity,STIG ID,Rule Title,Fix Text,Discussion,Check Content,Notes", ",")

    Dim colRelease As Integer
    Dim colGroup As Integer
    Dim colSev As Integer
    Dim colStig As Integer
    Dim colTitle As Integer
    Dim colFix As Integer
    Dim colDisc As Integer
    Dim colCheck As Integer

    colRelease = RD_FindCol(wsSrc, "Release Info")
    colGroup = RD_FindCol(wsSrc, "Group ID")
    colSev = RD_FindCol(wsSrc, "Severity")
    colStig = RD_FindCol(wsSrc, "STIG ID")
    colTitle = RD_FindCol(wsSrc, "Rule Title")
    colFix = RD_FindCol(wsSrc, "Fix Text")
    colDisc = RD_FindCol(wsSrc, "Discussion")
    colCheck = RD_FindCol(wsSrc, "Check Content")

    If colRelease = 0 Or colGroup = 0 Or colSev = 0 Or colStig = 0 Or colTitle = 0 _
        Or colFix = 0 Or colDisc = 0 Or colCheck = 0 Then
        MsgBox "One or more expected headers were not found on '" & SRC_SHEET_NAME & "'. Check row 1 spelling.", vbCritical
        GoTo CleanExit
    End If

    Dim lastRow As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, colGroup).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found on '" & SRC_SHEET_NAME & "'.", vbCritical
        GoTo CleanExit
    End If

    Dim h As Integer
    For h = LBound(headers) To UBound(headers)
        wsOut.Cells(1, h + 1).Value = headers(h)
    Next h
    wsOut.Rows(1).Font.Bold = True

    ' --- Discover the distinct Release Info values present ---
    Dim releaseKeys As New Collection
    Dim releaseDates As New Collection
    Dim r As Long
    Dim relTxt As String
    Dim relDate As Date

    For r = 2 To lastRow
        relTxt = Trim(CStr(wsSrc.Cells(r, colRelease).Value))
        If relTxt <> "" Then
            If Not RD_KeyExists(releaseKeys, relTxt) Then
                relDate = RD_ParseBenchmarkDate(relTxt)
                releaseKeys.Add relTxt, relTxt
                releaseDates.Add relDate, relTxt
            End If
        End If
    Next r

    Dim rk As Variant
    Dim msg As String
    If releaseKeys.Count <> 2 Then
        msg = "Expected exactly 2 distinct 'Release Info' values, found " & releaseKeys.Count & ":" & vbCrLf
        For Each rk In releaseKeys
            msg = msg & "  - " & rk & vbCrLf
        Next rk
        MsgBox msg, vbCritical
        GoTo CleanExit
    End If

    Dim relArr(1 To 2) As String
    Dim dateArr(1 To 2) As Date
    Dim idx As Integer
    idx = 1
    For Each rk In releaseKeys
        relArr(idx) = CStr(rk)
        dateArr(idx) = releaseDates(CStr(rk))
        idx = idx + 1
    Next rk

    If dateArr(1) <= dateArr(2) Then
        oldRelease = relArr(1)
        newRelease = relArr(2)
    Else
        oldRelease = relArr(2)
        newRelease = relArr(1)
    End If

    ' --- Index rows by revision + Group ID ---
    Dim oldRows As New Collection
    Dim newRows As New Collection
    Dim gTxt As String

    For r = 2 To lastRow
        relTxt = Trim(CStr(wsSrc.Cells(r, colRelease).Value))
        gTxt = Trim(CStr(wsSrc.Cells(r, colGroup).Value))
        If gTxt <> "" Then
            If relTxt = oldRelease Then
                If Not RD_KeyExists(oldRows, gTxt) Then oldRows.Add r, gTxt
            ElseIf relTxt = newRelease Then
                If Not RD_KeyExists(newRows, gTxt) Then newRows.Add r, gTxt
            End If
        End If
    Next r

    ' --- Copy every row as-is, then fill in Notes ---
    Dim outRow As Long
    Dim noteTxt As String
    Dim nr As Long
    outRow = 2

    For r = 2 To lastRow
        wsOut.Cells(outRow, 1).Value = wsSrc.Cells(r, colRelease).Value
        wsOut.Cells(outRow, 2).Value = wsSrc.Cells(r, colGroup).Value
        wsOut.Cells(outRow, 3).Value = wsSrc.Cells(r, colSev).Value
        wsOut.Cells(outRow, 4).Value = wsSrc.Cells(r, colStig).Value
        wsOut.Cells(outRow, 5).Value = wsSrc.Cells(r, colTitle).Value
        wsOut.Cells(outRow, 6).Value = wsSrc.Cells(r, colFix).Value
        wsOut.Cells(outRow, 7).Value = wsSrc.Cells(r, colDisc).Value
        wsOut.Cells(outRow, 8).Value = wsSrc.Cells(r, colCheck).Value

        relTxt = Trim(CStr(wsSrc.Cells(r, colRelease).Value))
        gTxt = Trim(CStr(wsSrc.Cells(r, colGroup).Value))
        noteTxt = ""

        If relTxt = oldRelease And gTxt <> "" Then
            If Not RD_KeyExists(newRows, gTxt) Then
                noteTxt = "Deleted - not present in " & newRelease
            Else
                nr = newRows(gTxt)
                noteTxt = RD_ClassifyChange(wsSrc, r, nr, colSev, colStig, colTitle, colFix, colDisc, colCheck)
            End If
        ElseIf relTxt = newRelease And gTxt <> "" Then
            If Not RD_KeyExists(oldRows, gTxt) Then
                noteTxt = "New - not present in " & oldRelease
            End If
        End If

        wsOut.Cells(outRow, 9).Value = noteTxt
        RD_ColorNote wsOut.Cells(outRow, 9), noteTxt

        outRow = outRow + 1
    Next r

    wsOut.Range("A1:I1").AutoFilter
    wsOut.Columns("A:I").AutoFit

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Release diff complete." & vbCrLf & "Older: " & oldRelease & vbCrLf & "Newer: " & newRelease, vbInformation
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Function RD_GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = sheetName
    End If
    Set RD_GetOrCreateSheet = ws
End Function

Function RD_FindCol(ws As Worksheet, headerName As String) As Integer
    Dim c As Integer
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(1, c).Value)) = headerName Then
            RD_FindCol = c
            Exit Function
        End If
    Next c
    RD_FindCol = 0
End Function

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

Function RD_ClassifyChange(ws As Worksheet, oldRow As Long, newRow As Long, _
                            colSev As Integer, colStig As Integer, colTitle As Integer, _
                            colFix As Integer, colDisc As Integer, colCheck As Integer) As String
    Dim sevA As String, sevB As String
    Dim stigA As String, stigB As String
    Dim titleA As String, titleB As String
    Dim fixA As String, fixB As String
    Dim discA As String, discB As String
    Dim checkA As String, checkB As String
    Dim changedFields As String

    sevA = RD_SafeStr(ws.Cells(oldRow, colSev).Value)
    sevB = RD_SafeStr(ws.Cells(newRow, colSev).Value)
    stigA = RD_SafeStr(ws.Cells(oldRow, colStig).Value)
    stigB = RD_SafeStr(ws.Cells(newRow, colStig).Value)
    titleA = RD_SafeStr(ws.Cells(oldRow, colTitle).Value)
    titleB = RD_SafeStr(ws.Cells(newRow, colTitle).Value)
    fixA = RD_SafeStr(ws.Cells(oldRow, colFix).Value)
    fixB = RD_SafeStr(ws.Cells(newRow, colFix).Value)
    discA = RD_SafeStr(ws.Cells(oldRow, colDisc).Value)
    discB = RD_SafeStr(ws.Cells(newRow, colDisc).Value)
    checkA = RD_SafeStr(ws.Cells(oldRow, colCheck).Value)
    checkB = RD_SafeStr(ws.Cells(newRow, colCheck).Value)

    changedFields = ""
    If sevA <> sevB Then changedFields = RD_AppendField(changedFields, "Severity")
    If stigA <> stigB Then changedFields = RD_AppendField(changedFields, "STIG ID")
    If titleA <> titleB Then changedFields = RD_AppendField(changedFields, "Rule Title")
    If fixA <> fixB Then changedFields = RD_AppendField(changedFields, "Fix Text")
    If discA <> discB Then changedFields = RD_AppendField(changedFields, "Discussion")
    If checkA <> checkB Then changedFields = RD_AppendField(changedFields, "Check Content")

    If changedFields = "" Then
        RD_ClassifyChange = "Unchanged"
        Exit Function
    End If

    ' A Severity or STIG ID change is never purely cosmetic
    If sevA <> sevB Or stigA <> stigB Then
        RD_ClassifyChange = "Modified - Need Review (" & changedFields & ")"
        Exit Function
    End If

    Dim onlyEditorial As Boolean
    onlyEditorial = True
    If RD_Normalize(titleA) <> RD_Normalize(titleB) Then onlyEditorial = False
    If RD_Normalize(fixA) <> RD_Normalize(fixB) Then onlyEditorial = False
    If RD_Normalize(discA) <> RD_Normalize(discB) Then onlyEditorial = False
    If RD_Normalize(checkA) <> RD_Normalize(checkB) Then onlyEditorial = False

    If onlyEditorial Then
        RD_ClassifyChange = "Modified - Editorial Only (" & changedFields & ")"
    Else
        RD_ClassifyChange = "Modified - Need Review (" & changedFields & ")"
    End If
End Function

Function RD_AppendField(existingList As String, fieldName As String) As String
    If existingList = "" Then
        RD_AppendField = fieldName
    Else
        RD_AppendField = existingList & ", " & fieldName
    End If
End Function

Function RD_Normalize(s As String) As String
    Dim t As String
    Dim punct As Variant
    Dim p As Variant

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
