Sub RunDiff()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsA As Worksheet, wsB As Worksheet, wsD As Worksheet
    Set wsA = Sheets("Rev5")
    Set wsB = Sheets("Rev8")
    Set wsD = Sheets("Diff")

    ' Clear old results
    Dim lastClearRow As Long
    lastClearRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    If lastClearRow < 2 Then lastClearRow = 2
    wsD.Range("A2:Z" & lastClearRow + 10).ClearContents
    wsD.Range("A2:Z" & lastClearRow + 10).Interior.ColorIndex = xlColorIndexNone

    Dim headers() As String
    headers = Split("Benchmark Name,Benchmark ID,Release Info,Version,Group ID,Severity,Rule ID,STIG ID,Classification,Asset Posture,SRG ID,Rule Title,Fix Text,Discussion,CCIs,Legacy IDs,Check Content,Check Content Ref,IA Controls,Weight,False Positives,False Negatives,Documentable", ",")
    Dim keyCol As String
    keyCol = "Group ID"

    Dim colA(0 To 30) As Integer, colB(0 To 30) As Integer
    Dim i As Integer
    For i = LBound(headers) To UBound(headers)
        colA(i) = FindCol(wsA, headers(i))
        colB(i) = FindCol(wsB, headers(i))
    Next i

    Dim keyIdx As Integer
    keyIdx = -1
    For i = LBound(headers) To UBound(headers)
        If headers(i) = keyCol Then keyIdx = i
    Next i
    If keyIdx = -1 Then
        MsgBox "'Group ID' not found in the header list inside the macro.", vbCritical
        GoTo CleanExit
    End If

    Dim keyColA As Integer, keyColB As Integer
    keyColA = colA(keyIdx)
    keyColB = colB(keyIdx)
    If keyColA = 0 Or keyColB = 0 Then
        MsgBox "Could not find a 'Group ID' column header on the Rev5 or Rev8 tab. Check spelling/spacing in row 1.", vbCritical
        GoTo CleanExit
    End If

    Dim lastA As Long, lastB As Long
    lastA = wsA.Cells(wsA.Rows.Count, keyColA).End(xlUp).Row
    lastB = wsB.Cells(wsB.Rows.Count, keyColB).End(xlUp).Row
    If lastA < 2 Then
        MsgBox "No data found on Rev5 (nothing below row 1).", vbCritical
        GoTo CleanExit
    End If
    If lastB < 2 Then
        MsgBox "No data found on Rev8 (nothing below row 1).", vbCritical
        GoTo CleanExit
    End If

    ' Row-lookup collections, keyed by Group ID text (cross-platform, no Scripting.Dictionary)
    Dim rowsA As New Collection, rowsB As New Collection
    Dim allKeys As New Collection
    Dim r As Long, kTxt As String

    For r = 2 To lastA
        kTxt = Trim(CStr(wsA.Cells(r, keyColA).Value))
        If kTxt <> "" Then
            If Not KeyExists(rowsA, kTxt) Then
                rowsA.Add r, kTxt
                allKeys.Add kTxt, kTxt
            End If
        End If
    Next r

    For r = 2 To lastB
        kTxt = Trim(CStr(wsB.Cells(r, keyColB).Value))
        If kTxt <> "" Then
            If Not KeyExists(rowsB, kTxt) Then rowsB.Add r, kTxt
            If Not KeyExists(allKeys, kTxt) Then allKeys.Add kTxt, kTxt
        End If
    Next r

    Dim outRow As Long
    outRow = 2

    Dim key As Variant
    For Each key In allKeys
        Dim inA As Boolean, inB As Boolean
        inA = KeyExists(rowsA, CStr(key))
        inB = KeyExists(rowsB, CStr(key))

        Dim rowA As Long, rowB As Long
        If inA Then rowA = rowsA(CStr(key))
        If inB Then rowB = rowsB(CStr(key))

        wsD.Cells(outRow, 2).Value = key

        Dim status As String
        If inA And Not inB Then
            status = "Removed"
        ElseIf inB And Not inA Then
            status = "Added"
        Else
            status = "Unchanged"
        End If

        Dim anyChanged As Boolean
        anyChanged = False

        Dim colOut As Integer
        colOut = 3
        For i = LBound(headers) To UBound(headers)
            If headers(i) <> keyCol Then
                Dim valA As String, valB As String
                valA = "": valB = ""
                If inA And colA(i) > 0 Then valA = SafeStr(wsA.Cells(rowA, colA(i)).Value)
                If inB And colB(i) > 0 Then valB = SafeStr(wsB.Cells(rowB, colB(i)).Value)

                If inA And inB Then
                    If valA <> valB Then
                        anyChanged = True
                        WriteWordDiff wsD.Cells(outRow, colOut), valA, valB
                    Else
                        wsD.Cells(outRow, colOut).Value = valA
                    End If
                ElseIf inA And Not inB Then
                    wsD.Cells(outRow, colOut).Value = valA
                ElseIf inB And Not inA Then
                    wsD.Cells(outRow, colOut).Value = valB
                End If
            End If
            colOut = colOut + 1
        Next i

        If inA And inB And anyChanged Then status = "Modified"
        wsD.Cells(outRow, 1).Value = status

        Dim fillColor As Long
        Select Case status
            Case "Added": fillColor = RGB(198, 239, 206)
            Case "Removed": fillColor = RGB(255, 199, 206)
            Case "Modified": fillColor = RGB(255, 235, 156)
            Case Else: fillColor = -1
        End Select
        If fillColor <> -1 Then
            wsD.Range(wsD.Cells(outRow, 1), wsD.Cells(outRow, colOut - 1)).Interior.Color = fillColor
        End If

        outRow = outRow + 1
    Next key

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Diff complete: " & (outRow - 2) & " rows processed.", vbInformation
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description & " (line context: row " & outRow & ")", vbCritical
End Sub

Function KeyExists(col As Collection, key As String) As Boolean
    Dim x As Variant
    On Error GoTo NotFound
    x = col(key)
    KeyExists = True
    Exit Function
NotFound:
    KeyExists = False
End Function

Function SafeStr(v As Variant) As String
    If IsError(v) Then
        SafeStr = ""
    ElseIf IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = CStr(v)
    End If
End Function

Function FindCol(ws As Worksheet, headerName As String) As Integer
    Dim c As Integer, lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(1, c).Value)) = headerName Then
            FindCol = c
            Exit Function
        End If
    Next c
    FindCol = 0
End Function

Sub WriteWordDiff(targetCell As Range, oldText As String, newText As String)
    Dim oldWords() As String, newWords() As String
    oldWords = Split(Trim(oldText), " ")
    newWords = Split(Trim(newText), " ")

    Dim ops() As Variant
    ops = WordDiffOps(oldWords, newWords)

    Dim fullText As String
    fullText = ""
    Dim n As Long
    For n = LBound(ops) To UBound(ops)
        fullText = fullText & ops(n)(1) & " "
    Next n
    If Len(fullText) = 0 Then fullText = " "

    targetCell.Value = fullText

    Dim pos As Long
    pos = 1
    For n = LBound(ops) To UBound(ops)
        Dim tag As String, word As String
        tag = ops(n)(0)
        word = ops(n)(1)
        Dim wlen As Long
        wlen = Len(word)
        If wlen > 0 Then
            With targetCell.Characters(pos, wlen).Font
                If tag = "del" Then
                    .Color = RGB(156, 0, 6)
                    .Strikethrough = True
                ElseIf tag = "ins" Then
                    .Color = RGB(0, 97, 0)
                    .Underline = xlUnderlineStyleSingle
                Else
                    .Color = RGB(0, 0, 0)
                    .Strikethrough = False
                    .Underline = xlUnderlineStyleNone
                End If
            End With
        End If
        pos = pos + wlen + 1
    Next n
End Sub

Function WordDiffOps(a() As String, b() As String) As Variant
    Dim la As Long, lb As Long
    Dim results As Collection
    Dim prefixLen As Long, suffixLen As Long
    Dim kk As Long
    Dim midALen As Long, midBLen As Long
    Dim midA() As String, midB() As String
    Dim midOps As Collection
    Dim item As Variant

    la = UBound(a) + 1
    lb = UBound(b) + 1
    Set results = New Collection

    If la = 0 And lb = 0 Then
        WordDiffOps = CollectionToArray(results)
        Exit Function
    End If

    prefixLen = 0
    Do While prefixLen < la And prefixLen < lb
        If a(prefixLen) = b(prefixLen) Then
            prefixLen = prefixLen + 1
        Else
            Exit Do
        End If
    Loop

    suffixLen = 0
    Do While suffixLen < (la - prefixLen) And suffixLen < (lb - prefixLen)
        If a(la - 1 - suffixLen) = b(lb - 1 - suffixLen) Then
            suffixLen = suffixLen + 1
        Else
            Exit Do
        End If
    Loop

    For kk = 0 To prefixLen - 1
        results.Add Array("eq", a(kk))
    Next kk

    midALen = la - prefixLen - suffixLen
    midBLen = lb - prefixLen - suffixLen

    If midALen > 0 Or midBLen > 0 Then
        If midALen > 0 Then
            ReDim midA(0 To midALen - 1)
            For kk = 0 To midALen - 1
                midA(kk) = a(prefixLen + kk)
            Next kk
        Else
            ReDim midA(0 To -1)
        End If

        If midBLen > 0 Then
            ReDim midB(0 To midBLen - 1)
            For kk = 0 To midBLen - 1
                midB(kk) = b(prefixLen + kk)
            Next kk
        Else
            ReDim midB(0 To -1)
        End If

        Set midOps = LCSDiff(midA, midB)
        For Each item In midOps
            results.Add item
        Next item
    End If

    For kk = 0 To suffixLen - 1
        results.Add Array("eq", a(la - suffixLen + kk))
    Next kk

    WordDiffOps = CollectionToArray(results)
End Function

Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    If col.Count = 0 Then
        ReDim arr(0 To 0)
        arr(0) = Array("eq", "")
        CollectionToArray = arr
        Exit Function
    End If
    ReDim arr(0 To col.Count - 1)
    Dim i As Long
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    CollectionToArray = arr
End Function

Function LCSDiff(a() As String, b() As String) As Collection
    Dim la As Long, lb As Long
    Dim result As Collection
    Dim dp() As Long
    Dim x As Long, y As Long
    Dim revOps As Collection
    Dim n As Long
    Dim j As Long, i2 As Long

    la = UBound(a) + 1
    lb = UBound(b) + 1
    If la < 0 Then la = 0
    If lb < 0 Then lb = 0

    Set result = New Collection

    If la = 0 And lb = 0 Then
        Set LCSDiff = result
        Exit Function
    End If
    If la = 0 Then
        For j = 0 To lb - 1
            result.Add Array("ins", b(j))
        Next j
        Set LCSDiff = result
        Exit Function
    End If
    If lb = 0 Then
        For i2 = 0 To la - 1
            result.Add Array("del", a(i2))
        Next i2
        Set LCSDiff = result
        Exit Function
    End If

    ReDim dp(0 To la, 0 To lb)
    For x = 1 To la
        For y = 1 To lb
            If a(x - 1) = b(y - 1) Then
                dp(x, y) = dp(x - 1, y - 1) + 1
            ElseIf dp(x - 1, y) >= dp(x, y - 1) Then
                dp(x, y) = dp(x - 1, y)
            Else
                dp(x, y) = dp(x, y - 1)
            End If
        Next y
    Next x

    Set revOps = New Collection
    x = la
    y = lb
    Do While x > 0 And y > 0
        If a(x - 1) = b(y - 1) Then
            revOps.Add Array("eq", a(x - 1))
            x = x - 1
            y = y - 1
        ElseIf dp(x - 1, y) >= dp(x, y - 1) Then
            revOps.Add Array("del", a(x - 1))
            x = x - 1
        Else
            revOps.Add Array("ins", b(y - 1))
            y = y - 1
        End If
    Loop
    Do While x > 0
        revOps.Add Array("del", a(x - 1))
        x = x - 1
    Loop
    Do While y > 0
        revOps.Add Array("ins", b(y - 1))
        y = y - 1
    Loop

    For n = revOps.Count To 1 Step -1
        result.Add revOps(n)
    Next n

    Set LCSDiff = result
End Function
