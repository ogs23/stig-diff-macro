Sub RunDiff()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsA As Worksheet, wsB As Worksheet, wsD As Worksheet
    Set wsA = Sheets("Rev5")
    Set wsB = Sheets("Rev8")
    Set wsD = Sheets("Diff")

    Dim lastClearRow As Long
    lastClearRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    If lastClearRow < 2 Then lastClearRow = 2
    wsD.Range("A2:Z" & lastClearRow + 10).ClearContents
    wsD.Range("A2:Z" & lastClearRow + 10).Interior.ColorIndex = xlColorIndexNone

    Dim headers() As String
    headers = Split("Benchmark Name,Benchmark ID,Release Info,Version,Group ID,Severity,Rule ID,STIG ID,Classification,Asset Posture,SRG ID,Rule Title,Fix Text,Discussion,CCIs,Legacy IDs,Check Content,Check Content Ref,IA Controls,Weight,False Positives,False Negatives,Documentable", ",")
    Dim keyCol As String
    keyCol = "Group ID"

    Dim colA(0 To 30) As Integer
    Dim colB(0 To 30) As Integer
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
        MsgBox "Group ID not found in header list.", vbCritical
        GoTo CleanExit
    End If

    Dim keyColA As Integer
    Dim keyColB As Integer
    keyColA = colA(keyIdx)
    keyColB = colB(keyIdx)
    If keyColA = 0 Or keyColB = 0 Then
        MsgBox "Could not find Group ID column on Rev5 or Rev8.", vbCritical
        GoTo CleanExit
    End If

    Dim lastA As Long
    Dim lastB As Long
    lastA = wsA.Cells(wsA.Rows.Count, keyColA).End(xlUp).Row
    lastB = wsB.Cells(wsB.Rows.Count, keyColB).End(xlUp).Row
    If lastA < 2 Then
        MsgBox "No data on Rev5.", vbCritical
        GoTo CleanExit
    End If
    If lastB < 2 Then
        MsgBox "No data on Rev8.", vbCritical
        GoTo CleanExit
    End If

    Dim rowsA As New Collection
    Dim rowsB As New Collection
    Dim allKeys As New Collection
    Dim r As Long
    Dim kTxt As String

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
        Dim inA As Boolean
        Dim inB As Boolean
        inA = KeyExists(rowsA, CStr(key))
        inB = KeyExists(rowsB, CStr(key))

        Dim rowA As Long
        Dim rowB As Long
        rowA = 0
        rowB = 0
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
                Dim valA As String
                Dim valB As String
                valA = ""
                valB = ""
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
        fillColor = -1
        Select Case status
            Case "Added"
                fillColor = RGB(198, 239, 206)
            Case "Removed"
                fillColor = RGB(255, 199, 206)
            Case "Modified"
                fillColor = RGB(255, 235, 156)
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
    MsgBox "Error: " & Err.Description & " (around output row " & outRow & ")", vbCritical
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
    Dim c As Integer
    Dim lastCol As Integer
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
    Dim oldWords() As String
    Dim newWords() As String
    oldWords = Split(Trim(oldText), " ")
    newWords = Split(Trim(newText), " ")

    Dim ops As Collection
    Set ops = LCSDiff(oldWords, newWords)

    Dim fullText As String
    fullText = ""
    Dim item As Variant
    For Each item In ops
        fullText = fullText & item(1) & " "
    Next item
    If Len(fullText) = 0 Then fullText = " "

    targetCell.Value = fullText

    Dim pos As Long
    pos = 1
    For Each item In ops
        Dim tag As String
        Dim word As String
        tag = item(0)
        word = item(1)
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
    Next item
End Sub

Function LCSDiff(a() As String, b() As String) As Collection
    Dim la As Long
    Dim lb As Long
    la = UBound(a) + 1
    lb = UBound(b) + 1
    If la < 0 Then la = 0
    If lb < 0 Then lb = 0

    Dim result As New Collection

    If la = 0 And lb = 0 Then
        Set LCSDiff = result
        Exit Function
    End If
    If la = 0 Then
        Dim j As Long
        For j = 0 To lb - 1
            result.Add Array("ins", b(j))
        Next j
        Set LCSDiff = result
        Exit Function
    End If
    If lb = 0 Then
        Dim i2 As Long
        For i2 = 0 To la - 1
            result.Add Array("del", a(i2))
        Next i2
        Set LCSDiff = result
        Exit Function
    End If

    Dim dp() As Long
    ReDim dp(0 To la, 0 To lb)
    Dim x As Long
    Dim y As Long
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

    Dim revOps As New Collection
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

    Dim n As Long
    For n = revOps.Count To 1 Step -1
        result.Add revOps(n)
    Next n

    Set LCSDiff = result
End Function
