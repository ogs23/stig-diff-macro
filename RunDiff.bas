Sub RunDiff()
    Application.ScreenUpdating = False

    Dim wsA As Worksheet, wsB As Worksheet, wsD As Worksheet
    Set wsA = Sheets("Rev5")
    Set wsB = Sheets("Rev8")
    Set wsD = Sheets("Diff")

    wsD.Rows("2:100000").ClearContents
    wsD.Rows("2:100000").Interior.ColorIndex = xlColorIndexNone

    Dim headers() As String
    headers = Split("Benchmark Name,Benchmark ID,Release Info,Version,Group ID,Severity,Rule ID,STIG ID,Classification,Asset Posture,SRG ID,Rule Title,Fix Text,Discussion,CCIs,Legacy IDs,Check Content,Check Content Ref,IA Controls,Weight,False Positives,False Negatives,Documentable", ",")

    Dim keyCol As String
    keyCol = "Group ID"

    Dim colA() As Integer, colB() As Integer
    ReDim colA(LBound(headers) To UBound(headers))
    ReDim colB(LBound(headers) To UBound(headers))

    Dim i As Integer
    For i = LBound(headers) To UBound(headers)
        colA(i) = FindCol(wsA, headers(i))
        colB(i) = FindCol(wsB, headers(i))
    Next i

    Dim keyIdxInHeaders As Integer
    For i = LBound(headers) To UBound(headers)
        If headers(i) = keyCol Then keyIdxInHeaders = i
    Next i

    Dim keyColA As Integer, keyColB As Integer
    keyColA = colA(keyIdxInHeaders)
    keyColB = colB(keyIdxInHeaders)

    If keyColA = 0 Or keyColB = 0 Then
        MsgBox "Could not find 'Group ID' column header on Rev5 or Rev8. Check spelling.", vbCritical
        Exit Sub
    End If

    Dim lastA As Long, lastB As Long
    lastA = wsA.Cells(wsA.Rows.Count, keyColA).End(xlUp).Row
    lastB = wsB.Cells(wsB.Rows.Count, keyColB).End(xlUp).Row

    Dim dictA As Object, dictB As Object
    Set dictA = CreateObject("Scripting.Dictionary")
    Set dictB = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 2 To lastA
        Dim kA As String
        kA = Trim(CStr(wsA.Cells(r, keyColA).Value))
        If kA <> "" Then
            If Not dictA.Exists(kA) Then dictA.Add kA, r
        End If
    Next r
    For r = 2 To lastB
        Dim kB As String
        kB = Trim(CStr(wsB.Cells(r, keyColB).Value))
        If kB <> "" Then
            If Not dictB.Exists(kB) Then dictB.Add kB, r
        End If
    Next r

    Dim allKeys As Collection
    Set allKeys = New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In dictA.Keys
        allKeys.Add k
        seen.Add k, True
    Next k
    For Each k In dictB.Keys
        If Not seen.Exists(k) Then
            allKeys.Add k
            seen.Add k, True
        End If
    Next k

    Dim outRow As Long
    outRow = 2

    Dim idx As Long
    For idx = 1 To allKeys.Count
        Dim key As String
        key = allKeys(idx)
        Dim inA As Boolean, inB As Boolean
        inA = dictA.Exists(key)
        inB = dictB.Exists(key)

        Dim rowA As Long, rowB As Long
        If inA Then rowA = dictA(key)
        If inB Then rowB = dictB(key)

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
                valA = ""
                valB = ""
                If inA And colA(i) > 0 Then valA = CStr(wsA.Cells(rowA, colA(i)).Value)
                If inB And colB(i) > 0 Then valB = CStr(wsB.Cells(rowB, colB(i)).Value)

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
        End Select
        If status <> "Unchanged" Then
            wsD.Range(wsD.Cells(outRow, 1), wsD.Cells(outRow, colOut - 1)).Interior.Color = fillColor
        End If

        outRow = outRow + 1
    Next idx

    Application.ScreenUpdating = True
    MsgBox "Diff complete: " & (outRow - 2) & " rows processed.", vbInformation
End Sub

Function FindCol(ws As Worksheet, headerName As String) As Integer
    Dim c As Integer
    Dim lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(ws.Cells(1, c).Value) = headerName Then
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

    Dim ops As Collection
    Set ops = WordDiffOps(oldWords, newWords)

    Dim fullText As String
    fullText = ""
    Dim opItem As Variant
    For Each opItem In ops
        fullText = fullText & opItem(1) & " "
    Next opItem

    targetCell.Value = fullText

    Dim pos As Long
    pos = 1
    For Each opItem In ops
        Dim tag As String, word As String
        tag = opItem(0)
        word = opItem(1)
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
    Next opItem
End Sub

Function WordDiffOps(a() As String, b() As String) As Collection
    Dim la As Long, lb As Long
    la = UBound(a) - LBound(a) + 1
    lb = UBound(b) - LBound(b) + 1

    Dim results As Collection
    Set results = New Collection

    If la = 0 And lb = 0 Then
        Set WordDiffOps = results
        Exit Function
    End If

    Dim prefixLen As Long
    prefixLen = 0
    Do While prefixLen < la And prefixLen < lb
        If a(LBound(a) + prefixLen) = b(LBound(b) + prefixLen) Then
            prefixLen = prefixLen + 1
        Else
            Exit Do
        End If
    Loop

    Dim suffixLen As Long
    suffixLen = 0
    Do While suffixLen < (la - prefixLen) And suffixLen < (lb - prefixLen)
        If a(UBound(a) - suffixLen) = b(UBound(b) - suffixLen) Then
            suffixLen = suffixLen + 1
        Else
            Exit Do
        End If
    Loop

    Dim kk As Long
    For kk = 0 To prefixLen - 1
        results.Add Array("eq", a(LBound(a) + kk))
    Next kk

    Dim midALen As Long, midBLen As Long
    midALen = la - prefixLen - suffixLen
    midBLen = lb - prefixLen - suffixLen

    Dim midA As Collection, midB As Collection
    Set midA = New Collection
    Set midB = New Collection
    For kk = 0 To midALen - 1
        midA.Add a(LBound(a) + prefixLen + kk)
    Next kk
    For kk = 0 To midBLen - 1
        midB.Add b(LBound(b) + prefixLen + kk)
    Next kk

    Dim midOps As Collection
    Set midOps = LCSDiff(midA, midB)
    Dim item As Variant
    For Each item In midOps
        results.Add item
    Next item

    For kk = 0 To suffixLen - 1
        results.Add Array("eq", a(UBound(a) - suffixLen + 1 + kk))
    Next kk

    Set WordDiffOps = results
End Function

Function LCSDiff(a As Collection, b As Collection) As Collection
    Dim la As Long, lb As Long
    la = a.Count
    lb = b.Count

    Dim result As Collection
    Set result = New Collection

    If la = 0 And lb = 0 Then
        Set LCSDiff = result
        Exit Function
    End If
    If la = 0 Then
        Dim j As Long
        For j = 1 To lb
            result.Add Array("ins", b(j))
        Next j
        Set LCSDiff = result
        Exit Function
    End If
    If lb = 0 Then
        Dim i2 As Long
        For i2 = 1 To la
            result.Add Array("del", a(i2))
        Next i2
        Set LCSDiff = result
        Exit Function
    End If

    Dim dp() As Long
    ReDim dp(0 To la, 0 To lb)
    Dim x As Long, y As Long
    For x = 1 To la
        For y = 1 To lb
            If a(x) = b(y) Then
                dp(x, y) = dp(x - 1, y - 1) + 1
            Else
                If dp(x - 1, y) >= dp(x, y - 1) Then
                    dp(x, y) = dp(x - 1, y)
                Else
                    dp(x, y) = dp(x, y - 1)
                End If
            End If
        Next y
    Next x

    Dim revOps As Collection
    Set revOps = New Collection
    x = la: y = lb
    Do While x > 0 And y > 0
        If a(x) = b(y) Then
            revOps.Add Array("eq", a(x))
            x = x - 1: y = y - 1
        ElseIf dp(x - 1, y) >= dp(x, y - 1) Then
            revOps.Add Array("del", a(x))
            x = x - 1
        Else
            revOps.Add Array("ins", b(y))
            y = y - 1
        End If
    Loop
    Do While x > 0
        revOps.Add Array("del", a(x))
        x = x - 1
    Loop
    Do While y > 0
        revOps.Add Array("ins", b(y))
        y = y - 1
    Loop

    Dim n As Long
    For n = revOps.Count To 1 Step -1
        result.Add revOps(n)
    Next n

    Set LCSDiff = result
End Function
