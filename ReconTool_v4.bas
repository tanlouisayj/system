Attribute VB_Name = "ReconTool_v4"
' ================================================================
'  RECONCILIATION TOOL v4  —  Finance Team Macro
'
'  HOW TO INSTALL:
'    1. Press Alt+F11 in Excel to open the VBA editor
'    2. Go to File > Import File > select this .bas file
'       OR: Insert > Module, then paste everything below
'    3. Back in Excel: Developer > Macros > RunRecon > Run
'
'  FEATURES:
'    - Multi-column key matching (e.g. Cost Centre + Period)
'    - Smart date normalisation (022026 matches feb2026)
'    - Aggregation: Sum, Count, Average, Min, Max
'    - Filters on each file before aggregating
'    - Drill-down: separate sheet shows source rows per key
'    - Colour-coded output with summary block
' ================================================================

Option Explicit

' ---- CONFIG: change these defaults if your team always uses the same columns ----
Private Const DEFAULT_TOLERANCE As Double = 0.01
Private Const MAX_KEY_COLS As Integer = 5

' ================================================================
'  MAIN ENTRY POINT
' ================================================================
Public Sub RunRecon()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim f1Path As String, f2Path As String
    Dim wb1 As Workbook, wb2 As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet

    ' --- Pick File 1 ---
    MsgBox "Select File 1 (Source — e.g. Essbase export)", vbInformation, "Recon Tool v4"
    f1Path = PickFile()
    If f1Path = "" Then GoTo Cancelled

    ' --- Pick File 2 ---
    MsgBox "Select File 2 (Comparison — e.g. GL export)", vbInformation, "Recon Tool v4"
    f2Path = PickFile()
    If f2Path = "" Then GoTo Cancelled

    ' --- Open both files ---
    Set wb1 = Workbooks.Open(f1Path, ReadOnly:=True)
    Set ws1 = wb1.Sheets(1)
    Set wb2 = Workbooks.Open(f2Path, ReadOnly:=True)
    Set ws2 = wb2.Sheets(1)

    ' --- Show columns to user ---
    Dim h1 As String, h2 As String
    h1 = GetHeaders(ws1)
    h2 = GetHeaders(ws2)
    MsgBox "File 1 columns:" & vbNewLine & h1, vbInformation, "File 1 columns"
    MsgBox "File 2 columns:" & vbNewLine & h2, vbInformation, "File 2 columns"

    ' --- How many key columns? ---
    Dim numKeys As Integer
    numKeys = CInt(InputBox("How many key columns to match on? (1-5)" & vbNewLine & "e.g. 2 if matching on Cost Centre + Period", "Number of key columns", "1"))
    If numKeys < 1 Then numKeys = 1
    If numKeys > MAX_KEY_COLS Then numKeys = MAX_KEY_COLS

    ' --- Collect key column pairs ---
    Dim keyMaps(1 To 5, 1 To 3) As String  ' (i, 1)=col1, (i, 2)=col2, (i, 3)=normType
    Dim i As Integer
    For i = 1 To numKeys
        Dim hint As String
        hint = SuggestMatch(GetHeaderArray(ws1), GetHeaderArray(ws2), i)
        keyMaps(i, 1) = InputBox("Key column " & i & " — File 1 column name:" & vbNewLine & h1, "Key Col " & i & " — File 1", hint)
        If keyMaps(i, 1) = "" Then GoTo CleanUp
        keyMaps(i, 2) = InputBox("Key column " & i & " — File 2 column name:" & vbNewLine & h2, "Key Col " & i & " — File 2", keyMaps(i, 1))
        If keyMaps(i, 2) = "" Then GoTo CleanUp
        keyMaps(i, 3) = InputBox("Normalise column " & i & " as:" & vbNewLine & "  text   — trim whitespace" & vbNewLine & "  number — numeric comparison (001 = 1)" & vbNewLine & "  integer — strip decimals" & vbNewLine & "  date   — parse to MMM-YY (Jan-25)" & vbNewLine & "  lower  — lowercase" & vbNewLine & "  upper  — uppercase", "Normalise Key " & i, GuessNormType(keyMaps(i, 1)))
        If keyMaps(i, 3) = "" Then keyMaps(i, 3) = "text"
    Next i

    ' --- Amount columns & aggregation ---
    Dim amtCol1 As String, amtCol2 As String
    Dim aggType1 As String, aggType2 As String

    amtCol1 = InputBox("File 1 — amount/value column name:", "Amount Column — File 1")
    If amtCol1 = "" Then GoTo CleanUp
    aggType1 = InputBox("File 1 — aggregation type:" & vbNewLine & "  sum, count, average, min, max", "Aggregation — File 1", "sum")
    If aggType1 = "" Then aggType1 = "sum"

    amtCol2 = InputBox("File 2 — amount/value column name:", "Amount Column — File 2")
    If amtCol2 = "" Then GoTo CleanUp
    aggType2 = InputBox("File 2 — aggregation type:" & vbNewLine & "  sum, count, average, min, max", "Aggregation — File 2", aggType1)
    If aggType2 = "" Then aggType2 = "sum"

    ' --- Optional filters ---
    Dim filt1 As String, filt2 As String
    filt1 = InputBox("File 1 — filter rows? (optional)" & vbNewLine & "Format: ColumnName=Value" & vbNewLine & "e.g. Version=Actual   or leave blank for no filter", "Filter — File 1", "")
    filt2 = InputBox("File 2 — filter rows? (optional)" & vbNewLine & "Format: ColumnName=Value", "Filter — File 2", "")

    ' --- Labels & tolerance ---
    Dim lbl1 As String, lbl2 As String
    lbl1 = InputBox("Label for File 1 (e.g. Essbase, GL, Budget):", "Label", "File 1")
    lbl2 = InputBox("Label for File 2:", "Label", "File 2")
    Dim tol As Double
    tol = CDbl(InputBox("Tolerance (ignore differences smaller than):", "Tolerance", CStr(DEFAULT_TOLERANCE)))

    ' --- Run the reconciliation ---
    Call DoRecon(ws1, ws2, numKeys, keyMaps, amtCol1, amtCol2, aggType1, aggType2, filt1, filt2, lbl1, lbl2, tol)

CleanUp:
    wb1.Close SaveChanges:=False
    wb2.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

Cancelled:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Cancelled.", vbInformation
End Sub


' ================================================================
'  CORE RECON LOGIC
' ================================================================
Private Sub DoRecon(ws1 As Worksheet, ws2 As Worksheet, _
                    numKeys As Integer, keyMaps() As String, _
                    amtCol1 As String, amtCol2 As String, _
                    aggType1 As String, aggType2 As String, _
                    filt1 As String, filt2 As String, _
                    lbl1 As String, lbl2 As String, _
                    tol As Double)

    ' --- Find column indices ---
    Dim kIdx1(1 To 5) As Long, kIdx2(1 To 5) As Long
    Dim aIdx1 As Long, aIdx2 As Long
    Dim i As Integer

    For i = 1 To numKeys
        kIdx1(i) = FindColIndex(ws1, keyMaps(i, 1))
        kIdx2(i) = FindColIndex(ws2, keyMaps(i, 2))
        If kIdx1(i) = 0 Then MsgBox "Column '" & keyMaps(i, 1) & "' not found in File 1.": Exit Sub
        If kIdx2(i) = 0 Then MsgBox "Column '" & keyMaps(i, 2) & "' not found in File 2.": Exit Sub
    Next i

    aIdx1 = FindColIndex(ws1, amtCol1)
    aIdx2 = FindColIndex(ws2, amtCol2)
    If aIdx1 = 0 Then MsgBox "Column '" & amtCol1 & "' not found in File 1.": Exit Sub
    If aIdx2 = 0 Then MsgBox "Column '" & amtCol2 & "' not found in File 2.": Exit Sub

    ' --- Parse filters ---
    Dim fCol1 As String, fVal1 As String
    Dim fCol2 As String, fVal2 As String
    Dim fIdx1 As Long, fIdx2 As Long
    ParseFilter filt1, fCol1, fVal1
    ParseFilter filt2, fCol2, fVal2
    If fCol1 <> "" Then fIdx1 = FindColIndex(ws1, fCol1)
    If fCol2 <> "" Then fIdx2 = FindColIndex(ws2, fCol2)

    ' --- Build aggregation dictionaries ---
    Dim dict1 As Object, dict2 As Object, cntDict1 As Object, cntDict2 As Object
    Set dict1 = CreateObject("Scripting.Dictionary")
    Set dict2 = CreateObject("Scripting.Dictionary")
    Set cntDict1 = CreateObject("Scripting.Dictionary")
    Set cntDict2 = CreateObject("Scripting.Dictionary")

    ' Store source rows for drill-down
    Dim srcDict1 As Object, srcDict2 As Object
    Set srcDict1 = CreateObject("Scripting.Dictionary")
    Set srcDict2 = CreateObject("Scripting.Dictionary")

    Dim lastRow1 As Long, lastRow2 As Long
    lastRow1 = ws1.Cells(ws1.Rows.Count, kIdx1(1)).End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, kIdx2(1)).End(xlUp).Row

    ' Aggregate File 1
    Dim r As Long
    For r = 2 To lastRow1
        If fIdx1 > 0 Then
            If LCase(Trim(CStr(ws1.Cells(r, fIdx1).Value))) <> LCase(Trim(fVal1)) Then GoTo NextR1
        End If
        Dim key1 As String
        key1 = MakeKey(ws1, r, numKeys, kIdx1, keyMaps, 1)
        If key1 = "" Then GoTo NextR1
        Dim v1 As Double
        v1 = ToDouble(ws1.Cells(r, aIdx1).Value)
        AggregateDict dict1, cntDict1, key1, v1, aggType1
        ' Store row reference for drill-down (row number)
        If Not srcDict1.exists(key1) Then srcDict1(key1) = key1 & "|" & r Else srcDict1(key1) = srcDict1(key1) & "," & r
NextR1:
    Next r

    ' Aggregate File 2
    For r = 2 To lastRow2
        If fIdx2 > 0 Then
            If LCase(Trim(CStr(ws2.Cells(r, fIdx2).Value))) <> LCase(Trim(fVal2)) Then GoTo NextR2
        End If
        Dim key2 As String
        key2 = MakeKey(ws2, r, numKeys, kIdx2, keyMaps, 2)
        If key2 = "" Then GoTo NextR2
        Dim v2 As Double
        v2 = ToDouble(ws2.Cells(r, aIdx2).Value)
        AggregateDict dict2, cntDict2, key2, v2, aggType2
        If Not srcDict2.exists(key2) Then srcDict2(key2) = key2 & "|" & r Else srcDict2(key2) = srcDict2(key2) & "," & r
NextR2:
    Next r

    ' Finalise averages
    FinaliseAvg dict1, cntDict1, aggType1
    FinaliseAvg dict2, cntDict2, aggType2

    ' --- Create output workbook ---
    Dim wbOut As Workbook
    Dim wsOut As Worksheet, wsDrill As Worksheet
    Set wbOut = Workbooks.Add
    Set wsOut = wbOut.Sheets(1)
    wsOut.Name = "Recon Output"
    Set wsDrill = wbOut.Sheets.Add(After:=wsOut)
    wsDrill.Name = "Source Drill-Down"

    ' --- Write headers on output sheet ---
    Dim hdrRow As Long
    hdrRow = 8  ' leave room for summary block
    Dim col As Long
    col = 1
    For i = 1 To numKeys
        Dim hdr As String
        If keyMaps(i, 1) = keyMaps(i, 2) Then hdr = keyMaps(i, 1) Else hdr = keyMaps(i, 1) & " / " & keyMaps(i, 2)
        wsOut.Cells(hdrRow, col).Value = hdr
        col = col + 1
    Next i
    wsOut.Cells(hdrRow, col).Value = lbl1 & " (" & UCase(Left(aggType1, 1)) & Mid(aggType1, 2) & " of " & amtCol1 & ")"
    wsOut.Cells(hdrRow, col + 1).Value = lbl2 & " (" & UCase(Left(aggType2, 1)) & Mid(aggType2, 2) & " of " & amtCol2 & ")"
    wsOut.Cells(hdrRow, col + 2).Value = "Difference"
    wsOut.Cells(hdrRow, col + 3).Value = "Status"
    wsOut.Cells(hdrRow, col + 4).Value = "Drill-Down Sheet Row"
    wsOut.Rows(hdrRow).Font.Bold = True
    wsOut.Rows(hdrRow).Interior.Color = RGB(240, 240, 240)

    ' --- Write data rows ---
    Dim outRow As Long
    outRow = hdrRow + 1
    Dim cntMatch As Long, cntBreak As Long, cntMiss1 As Long, cntMiss2 As Long

    Dim usedKeys As Object
    Set usedKeys = CreateObject("Scripting.Dictionary")

    ' Write File 1 keys
    Dim k As Variant
    For Each k In dict1.Keys
        usedKeys(k) = 1
        Dim val1 As Double, val2b As Double, diff As Double
        val1 = dict1(k)
        col = 1
        ' Write key parts
        Dim parts() As String
        parts = Split(CStr(k), Chr(167))
        For i = 1 To numKeys
            If i <= UBound(parts) + 1 Then wsOut.Cells(outRow, col).Value = parts(i - 1)
            col = col + 1
        Next i
        wsOut.Cells(outRow, col).Value = val1

        If dict2.exists(k) Then
            val2b = dict2(k)
            diff = val1 - val2b
            wsOut.Cells(outRow, col + 1).Value = val2b
            wsOut.Cells(outRow, col + 2).Value = diff
            If Abs(diff) > tol Then
                wsOut.Cells(outRow, col + 3).Value = "BREAK"
                wsOut.Rows(outRow).Interior.Color = RGB(255, 235, 235)
                wsOut.Cells(outRow, col + 3).Font.Color = RGB(180, 30, 30)
                cntBreak = cntBreak + 1
            Else
                wsOut.Cells(outRow, col + 3).Value = "Match"
                wsOut.Cells(outRow, col + 3).Font.Color = RGB(40, 110, 70)
                cntMatch = cntMatch + 1
            End If
        Else
            wsOut.Cells(outRow, col + 1).Value = ""
            wsOut.Cells(outRow, col + 2).Value = ""
            wsOut.Cells(outRow, col + 3).Value = "Not in " & lbl2
            wsOut.Rows(outRow).Interior.Color = RGB(255, 250, 230)
            wsOut.Cells(outRow, col + 3).Font.Color = RGB(150, 100, 0)
            cntMiss2 = cntMiss2 + 1
        End If
        outRow = outRow + 1
    Next k

    ' File 2 keys not in File 1
    For Each k In dict2.Keys
        If Not usedKeys.exists(k) Then
            parts = Split(CStr(k), Chr(167))
            col = 1
            For i = 1 To numKeys
                If i <= UBound(parts) + 1 Then wsOut.Cells(outRow, col).Value = parts(i - 1)
                col = col + 1
            Next i
            wsOut.Cells(outRow, col).Value = ""
            wsOut.Cells(outRow, col + 1).Value = dict2(k)
            wsOut.Cells(outRow, col + 2).Value = ""
            wsOut.Cells(outRow, col + 3).Value = "Not in " & lbl1
            wsOut.Rows(outRow).Interior.Color = RGB(255, 250, 230)
            wsOut.Cells(outRow, col + 3).Font.Color = RGB(150, 100, 0)
            cntMiss1 = cntMiss1 + 1
            outRow = outRow + 1
        End If
    Next k

    ' --- Summary block ---
    With wsOut
        .Range("A1:G7").Interior.Color = RGB(248, 248, 248)
        .Cells(1, 1).Value = "RECONCILIATION SUMMARY"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 13
        .Cells(2, 1).Value = lbl1 & " vs " & lbl2
        .Cells(2, 1).Font.Color = RGB(100, 100, 100)
        .Cells(4, 1).Value = "Matched": .Cells(4, 2).Value = cntMatch: .Cells(4, 2).Font.Color = RGB(40, 110, 70): .Cells(4, 2).Font.Bold = True
        .Cells(5, 1).Value = "Breaks": .Cells(5, 2).Value = cntBreak: .Cells(5, 2).Font.Color = RGB(180, 30, 30): .Cells(5, 2).Font.Bold = True
        .Cells(6, 1).Value = "Not in " & lbl2: .Cells(6, 2).Value = cntMiss2: .Cells(6, 2).Font.Color = RGB(150, 100, 0): .Cells(6, 2).Font.Bold = True
        .Cells(7, 1).Value = "Not in " & lbl1: .Cells(7, 2).Value = cntMiss1: .Cells(7, 2).Font.Color = RGB(150, 100, 0): .Cells(7, 2).Font.Bold = True
    End With

    ' --- Format amount columns ---
    Dim amtCols As String
    amtCols = "$" & ColLetter(numKeys + 1) & "$" & (hdrRow + 1) & ":$" & ColLetter(numKeys + 3) & "$" & (outRow - 1)
    wsOut.Range(amtCols).NumberFormat = "#,##0.00"
    wsOut.Columns.AutoFit
    wsOut.Rows(hdrRow).AutoFilter

    ' --- Build drill-down sheet ---
    BuildDrillSheet wsDrill, ws1, ws2, srcDict1, srcDict2, lbl1, lbl2, numKeys, keyMaps

    MsgBox "Done!" & vbNewLine & vbNewLine & _
           "Matched:         " & cntMatch & vbNewLine & _
           "Breaks:            " & cntBreak & vbNewLine & _
           "Not in " & lbl2 & ":   " & cntMiss2 & vbNewLine & _
           "Not in " & lbl1 & ":   " & cntMiss1 & vbNewLine & vbNewLine & _
           "See 'Source Drill-Down' tab to inspect raw rows per key.", _
           vbInformation, "Recon Complete"
End Sub


' ================================================================
'  DRILL-DOWN SHEET
' ================================================================
Private Sub BuildDrillSheet(wsDrill As Worksheet, ws1 As Worksheet, ws2 As Worksheet, _
                             srcDict1 As Object, srcDict2 As Object, _
                             lbl1 As String, lbl2 As String, _
                             numKeys As Integer, keyMaps() As String)
    Dim drillRow As Long
    drillRow = 1
    Dim lastCol1 As Long, lastCol2 As Long
    lastCol1 = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    lastCol2 = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column

    wsDrill.Cells(drillRow, 1).Value = "SOURCE DATA DRILL-DOWN"
    wsDrill.Cells(drillRow, 1).Font.Bold = True
    wsDrill.Cells(drillRow, 1).Font.Size = 12
    drillRow = drillRow + 2

    ' Combine all keys
    Dim allKeys As Object
    Set allKeys = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In srcDict1.Keys: allKeys(k) = 1: Next k
    For Each k In srcDict2.Keys: allKeys(k) = 1: Next k

    For Each k In allKeys.Keys
        ' Section header
        Dim parts() As String
        parts = Split(CStr(k), Chr(167))
        Dim keyDisplay As String
        keyDisplay = ""
        Dim i As Integer
        For i = 0 To UBound(parts)
            If i > 0 Then keyDisplay = keyDisplay & " | "
            If i < numKeys Then keyDisplay = keyDisplay & keyMaps(i + 1, 1) & ": " & parts(i) Else keyDisplay = keyDisplay & parts(i)
        Next i
        wsDrill.Cells(drillRow, 1).Value = keyDisplay
        wsDrill.Cells(drillRow, 1).Font.Bold = True
        wsDrill.Rows(drillRow).Interior.Color = RGB(230, 230, 230)
        drillRow = drillRow + 1

        ' File 1 rows
        If srcDict1.exists(k) Then
            wsDrill.Cells(drillRow, 1).Value = lbl1 & " source rows:"
            wsDrill.Cells(drillRow, 1).Font.Italic = True
            wsDrill.Cells(drillRow, 1).Font.Color = RGB(40, 110, 70)
            drillRow = drillRow + 1
            ' Write header
            Dim c As Long
            For c = 1 To lastCol1
                wsDrill.Cells(drillRow, c).Value = ws1.Cells(1, c).Value
                wsDrill.Cells(drillRow, c).Font.Bold = True
            Next c
            drillRow = drillRow + 1
            ' Write rows
            Dim rowNums1() As String
            rowNums1 = Split(Mid(CStr(srcDict1(k)), InStr(CStr(srcDict1(k)), "|") + 1), ",")
            Dim rn As Variant
            For Each rn In rowNums1
                For c = 1 To lastCol1
                    wsDrill.Cells(drillRow, c).Value = ws1.Cells(CLng(rn), c).Value
                Next c
                drillRow = drillRow + 1
            Next rn
        End If

        ' File 2 rows
        If srcDict2.exists(k) Then
            wsDrill.Cells(drillRow, 1).Value = lbl2 & " source rows:"
            wsDrill.Cells(drillRow, 1).Font.Italic = True
            wsDrill.Cells(drillRow, 1).Font.Color = RGB(30, 80, 160)
            drillRow = drillRow + 1
            For c = 1 To lastCol2
                wsDrill.Cells(drillRow, c).Value = ws2.Cells(1, c).Value
                wsDrill.Cells(drillRow, c).Font.Bold = True
            Next c
            drillRow = drillRow + 1
            Dim rowNums2() As String
            rowNums2 = Split(Mid(CStr(srcDict2(k)), InStr(CStr(srcDict2(k)), "|") + 1), ",")
            For Each rn In rowNums2
                For c = 1 To lastCol2
                    wsDrill.Cells(drillRow, c).Value = ws2.Cells(CLng(rn), c).Value
                Next c
                drillRow = drillRow + 1
            Next rn
        End If

        drillRow = drillRow + 1  ' blank line between keys
    Next k

    wsDrill.Columns.AutoFit
End Sub


' ================================================================
'  HELPERS
' ================================================================
Private Function MakeKey(ws As Worksheet, rowNum As Long, numKeys As Integer, _
                          kIdxArr() As Long, keyMaps() As String, side As Integer) As String
    Dim parts() As String
    ReDim parts(1 To numKeys)
    Dim i As Integer
    For i = 1 To numKeys
        parts(i) = NormaliseValue(CStr(ws.Cells(rowNum, kIdxArr(i)).Value), keyMaps(i, 3))
    Next i
    MakeKey = Join(parts, Chr(167))
End Function

Private Function NormaliseValue(val As String, normType As String) As String
    Dim s As String
    s = Trim(val)
    Select Case LCase(normType)
        Case "number", "num"
            NormaliseValue = CStr(CDbl(Replace(s, ",", "")) * 1)
        Case "integer", "int"
            NormaliseValue = CStr(CLng(Replace(s, ",", "")))
        Case "lower"
            NormaliseValue = LCase(s)
        Case "upper"
            NormaliseValue = UCase(s)
        Case "date"
            NormaliseValue = NormaliseDate(s)
        Case Else
            NormaliseValue = s
    End Select
End Function

Private Function NormaliseDate(s As String) As String
    On Error Resume Next
    ' Try parsing as actual date
    Dim d As Date
    d = CDate(s)
    If Err.Number = 0 Then
        NormaliseDate = Format(d, "mmm-yy")
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    ' Try MMYYYY or MMYY pattern
    Dim n As String
    n = s
    Dim alpha As String, nums As String
    Dim c As Integer
    For c = 1 To Len(n)
        Dim ch As String
        ch = Mid(n, c, 1)
        If ch >= "A" And ch <= "Z" Or ch >= "a" And ch <= "z" Then alpha = alpha & ch Else nums = nums & ch
    Next c
    Dim months(1 To 12) As String
    months(1) = "jan": months(2) = "feb": months(3) = "mar": months(4) = "apr"
    months(5) = "may": months(6) = "jun": months(7) = "jul": months(8) = "aug"
    months(9) = "sep": months(10) = "oct": months(11) = "nov": months(12) = "dec"
    If Len(alpha) >= 3 Then
        Dim mth As Integer
        For mth = 1 To 12
            If InStr(LCase(alpha), months(mth)) > 0 Then
                Dim yr As String
                yr = Right("20" & Right(nums, 2), 4)
                NormaliseDate = UCase(Left(months(mth), 1)) & Mid(months(mth), 2) & "-" & Right(yr, 2)
                Exit Function
            End If
        Next mth
    End If
    ' Try leading MMYYYY
    If Len(nums) = 6 Then
        Dim mm As Integer
        mm = CInt(Left(nums, 2))
        If mm >= 1 And mm <= 12 Then
            NormaliseDate = UCase(Left(months(mm), 1)) & Mid(months(mm), 2) & "-" & Right(nums, 2)
            Exit Function
        End If
    End If
    NormaliseDate = LCase(Trim(s))
End Function

Private Function GuessNormType(colName As String) As String
    Dim n As String
    n = LCase(colName)
    If InStr(n, "date") > 0 Or InStr(n, "period") > 0 Or InStr(n, "month") > 0 Or InStr(n, "mth") > 0 Then
        GuessNormType = "date": Exit Function
    End If
    ' Check if it looks like a date code (e.g. 022026)
    Dim nums As String, i As Integer
    For i = 1 To Len(colName)
        If Mid(colName, i, 1) >= "0" And Mid(colName, i, 1) <= "9" Then nums = nums & Mid(colName, i, 1)
    Next i
    If Len(nums) >= 4 Then GuessNormType = "date": Exit Function
    GuessNormType = "text"
End Function

Private Sub AggregateDict(dict As Object, cntDict As Object, key As String, val As Double, aggType As String)
    Select Case LCase(aggType)
        Case "count"
            If dict.exists(key) Then dict(key) = dict(key) + 1 Else dict(key) = 1
        Case "sum"
            If dict.exists(key) Then dict(key) = dict(key) + val Else dict(key) = val
        Case "average", "avg"
            If dict.exists(key) Then dict(key) = dict(key) + val Else dict(key) = val
            If cntDict.exists(key) Then cntDict(key) = cntDict(key) + 1 Else cntDict(key) = 1
        Case "min"
            If dict.exists(key) Then dict(key) = WorksheetFunction.Min(dict(key), val) Else dict(key) = val
        Case "max"
            If dict.exists(key) Then dict(key) = WorksheetFunction.Max(dict(key), val) Else dict(key) = val
        Case Else
            If dict.exists(key) Then dict(key) = dict(key) + val Else dict(key) = val
    End Select
End Sub

Private Sub FinaliseAvg(dict As Object, cntDict As Object, aggType As String)
    If LCase(aggType) <> "average" And LCase(aggType) <> "avg" Then Exit Sub
    Dim k As Variant
    For Each k In dict.Keys
        If cntDict.exists(k) And cntDict(k) > 0 Then dict(k) = dict(k) / cntDict(k)
    Next k
End Sub

Private Sub ParseFilter(filterStr As String, colName As String, colVal As String)
    If filterStr = "" Then colName = "": colVal = "": Exit Sub
    Dim pos As Integer
    pos = InStr(filterStr, "=")
    If pos > 0 Then
        colName = Trim(Left(filterStr, pos - 1))
        colVal = Trim(Mid(filterStr, pos + 1))
    End If
End Sub

Private Function SuggestMatch(headers1() As String, headers2() As String, idx As Integer) As String
    If idx <= UBound(headers1) + 1 Then
        Dim h As String
        h = headers1(idx - 1)
        Dim j As Integer
        For j = 0 To UBound(headers2)
            If LCase(Trim(headers2(j))) = LCase(Trim(h)) Then SuggestMatch = h: Exit Function
        Next j
        SuggestMatch = h
    End If
End Function

Private Function GetHeaders(ws As Worksheet) As String
    Dim arr() As String
    arr = GetHeaderArray(ws)
    GetHeaders = Join(arr, ", ")
End Function

Private Function GetHeaderArray(ws As Worksheet) As String()
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim arr() As String
    ReDim arr(0 To lastCol - 1)
    Dim j As Long
    For j = 1 To lastCol
        arr(j - 1) = CStr(ws.Cells(1, j).Value)
    Next j
    GetHeaderArray = arr
End Function

Private Function FindColIndex(ws As Worksheet, colName As String) As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim j As Long
    For j = 1 To lastCol
        If Trim(CStr(ws.Cells(1, j).Value)) = colName Then FindColIndex = j: Exit Function
    Next j
    FindColIndex = 0
End Function

Private Function ToDouble(val As Variant) As Double
    On Error Resume Next
    ToDouble = CDbl(Replace(CStr(val), ",", ""))
    On Error GoTo 0
End Function

Private Function PickFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "Excel & CSV", "*.xlsx;*.xls;*.csv"
    fd.Title = "Select file"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
End Function

Private Function ColLetter(n As Long) As String
    If n <= 26 Then ColLetter = Chr(64 + n) Else ColLetter = Chr(64 + Int((n - 1) / 26)) & Chr(65 + ((n - 1) Mod 26))
End Function
