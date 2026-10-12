Attribute VB_Name = "modWorkstation"
Option Explicit
' ---- Import settings ----
Private Const SOURCE_HAS_HEADER As Boolean = True
Private Const NUMBER_COLS As String = "5,9"   ' E and I are numbers
Private Const DATE_COLS As String = ""        ' e.g. "1,3" if A and C are dates
Private Const CSV_CHARSET As String = "utf-8" ' use "windows-1252" if accents look wrong
Private Const MARK_CLEAN As Long = 1   ' yellow
Private Const MARK_DUP As Long = 2     ' orange
' ===================== SETTINGS =====================
Private Const SHEET_DATA As String = "Workstation"
Private Const SHEET_REMOVED As String = "RemovedRows"
Private Const HEADER_ROW As Long = 1

Private Const COL_AMOUNT As Long = 5       ' E
Private Const COL_TYPE As Long = 6         ' F: OT, OS, OR, OM
Private Const COL_MERCHANT As Long = 7     ' G
Private Const COL_CHECK As Long = 9        ' I: rows equal to CLEAN_CHECK_VALUE are removed
Private Const CLEAN_CHECK_VALUE As Double = 50

' When a return (OR/OM) matches and both OT and OS remain, which is consumed?
Private Const RETURN_CONSUMES_OT_FIRST As Boolean = True

Private Const MAX_LIST_ROWS As Long = 5000

' ===================== HELPERS =====================
Private Function DataSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_DATA)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 513, "modWorkstation", "The sheet '" & SHEET_DATA & "' was not found."
    End If
    Set DataSheet = ws
End Function

Private Function LastDataRow(ByVal ws As Object) As Long
    Dim c As Object
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then
        LastDataRow = HEADER_ROW
    Else
        LastDataRow = c.Row
    End If
End Function

Private Function LastDataCol(ByVal ws As Object) As Long
    Dim c As Object, lc As Long
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If c Is Nothing Then lc = COL_CHECK Else lc = c.Column
    If lc < COL_CHECK Then lc = COL_CHECK
    LastDataCol = lc
End Function

' ===================== HIGHLIGHT HELPERS =====================
Private Function MarkColor(ByVal kind As Long) As Long
    If kind = MARK_CLEAN Then
        MarkColor = RGB(255, 255, 153)
    Else
        MarkColor = RGB(255, 192, 128)
    End If
End Function

' Returns, for each data row, 0 = not marked, MARK_CLEAN or MARK_DUP (by the fill color of column A)
Private Function ScanMarks(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As Byte()
    Dim k() As Byte, r As Long, clr As Long, cClean As Long, cDup As Long
    cClean = MarkColor(MARK_CLEAN)
    cDup = MarkColor(MARK_DUP)
    ReDim k(1 To lastRow - firstRow + 1)
    For r = firstRow To lastRow
        clr = ws.Cells(r, 1).Interior.Color
        If clr = cClean Then
            k(r - firstRow + 1) = MARK_CLEAN
        ElseIf clr = cDup Then
            k(r - firstRow + 1) = MARK_DUP
        End If
    Next r
    ScanMarks = k
End Function

' Paints (or clears) the rows flagged True, grouping consecutive rows into one range
Private Sub PaintRuns(ByVal ws As Worksheet, ByRef sel() As Boolean, ByVal firstRow As Long, _
                      ByVal lastCol As Long, ByVal clr As Long, ByVal clearIt As Boolean)
    Dim i As Long, n As Long, s As Long, rg As Range
    n = UBound(sel)
    i = 1
    Do While i <= n
        If sel(i) Then
            s = i
            Do While i < n
                If Not sel(i + 1) Then Exit Do
                i = i + 1
            Loop
            Set rg = ws.Range(ws.Cells(firstRow + s - 1, 1), ws.Cells(firstRow + i - 1, lastCol))
            If clearIt Then rg.Interior.Pattern = xlNone Else rg.Interior.Color = clr
        End If
        i = i + 1
    Loop
End Sub

' Removes old highlights of this kind, then highlights the new rows
Private Sub ApplyMarks(ByVal ws As Worksheet, ByRef newSel() As Boolean, ByVal kind As Long, _
                       ByVal firstRow As Long, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim kinds() As Byte, oldSel() As Boolean, i As Long, n As Long
    Dim prevSU As Boolean, errNum As Long, errDesc As String

    n = lastRow - firstRow + 1
    prevSU = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    kinds = ScanMarks(ws, firstRow, lastRow)
    ReDim oldSel(1 To n)
    For i = 1 To n
        oldSel(i) = (kinds(i) = kind)
    Next i
    PaintRuns ws, oldSel, firstRow, lastCol, 0, True
    PaintRuns ws, newSel, firstRow, lastCol, MarkColor(kind), False

Cleanup:
    Application.ScreenUpdating = prevSU
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "modWorkstation", errDesc
    Exit Sub

Fail:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Sub

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then
        CellText = ""
    Else
        CellText = CStr(v)
    End If
End Function

' Upper case, trimmed, no double spaces, no non-breaking spaces
Private Function NormText(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = Replace(CStr(v), Chr$(160), " ")
    NormText = UCase$(Application.WorksheetFunction.Trim(s))
End Function

Private Function TryAmount(ByVal v As Variant, ByRef result As Double) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(CStr(v)) = "" Then Exit Function
    End If
    If IsNumeric(v) Then
        result = CDbl(v)
        TryAmount = True
    End If
End Function

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Sub AddRow(ByVal d As Object, ByVal key As String, ByVal rowIdx As Long)
    If Not d.Exists(key) Then Set d(key) = New Collection
    d(key).Add rowIdx
End Sub

Private Function ListFor(ByVal d As Object, ByVal key As String) As Collection
    If d.Exists(key) Then
        Set ListFor = d(key)
    Else
        Set ListFor = New Collection
    End If
End Function

Private Sub AppendItems(ByVal dst As Collection, ByVal src As Collection, ByVal fromIdx As Long)
    Dim j As Long
    For j = fromIdx To src.Count
        dst.Add src(j)
    Next j
End Sub

' ===================== CLEAN (highlight only) =====================
Public Sub CleanData(ByVal lv As MSComctlLib.ListView)
    Dim ws As Worksheet, data As Variant
    Dim firstRow As Long, lastRow As Long, lastCol As Long
    Dim n As Long, i As Long, amt As Double
    Dim del() As Boolean
    Dim zeroCnt As Long, checkCnt As Long, total As Long

    On Error GoTo ErrHandler

    Set ws = DataSheet()
    firstRow = HEADER_ROW + 1
    lastRow = LastDataRow(ws)
    If lastRow < firstRow Then
        MsgBox "There is no data to clean.", vbInformation, "Clean"
        Exit Sub
    End If
    lastCol = LastDataCol(ws)

    data = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).Value
    n = UBound(data, 1)
    ReDim del(1 To n)

    For i = 1 To n
        If TryAmount(data(i, COL_AMOUNT), amt) Then
            If amt = 0 Then
                del(i) = True
                zeroCnt = zeroCnt + 1
            End If
        End If
        If Not del(i) Then
            If TryAmount(data(i, COL_CHECK), amt) Then
                If amt = CLEAN_CHECK_VALUE Then
                    del(i) = True
                    checkCnt = checkCnt + 1
                End If
            End If
        End If
    Next i

    total = zeroCnt + checkCnt
    ApplyMarks ws, del, MARK_CLEAN, firstRow, lastRow, lastCol   ' also clears stale Clean marks
    LogSecurityEvent gUsername, "Clean: highlighted " & total & " rows (E=0: " & zeroCnt & ", I=" & CLEAN_CHECK_VALUE & ": " & checkCnt & ")"
    LoadList lv

    If total = 0 Then
        MsgBox "Nothing to clean.", vbInformation, "Clean"
    Else
        MsgBox total & " rows highlighted (yellow):" & vbCrLf & _
               "   Amount (E) = 0:  " & zeroCnt & vbCrLf & _
               "   Column I = " & CLEAN_CHECK_VALUE & ":  " & checkCnt & vbCrLf & vbCrLf & _
               "Nothing was deleted. Review them, then press Delete.", vbInformation, "Clean"
    End If
    Exit Sub

ErrHandler:
    MsgBox "Clean error: " & Err.Number & " - " & Err.Description, vbExclamation, "Clean"
End Sub

' ===================== DUPLICATES (highlight only) =====================
Public Sub FindDuplicates(ByVal lv As MSComctlLib.ListView)
    Dim ws As Worksheet, data As Variant
    Dim firstRow As Long, lastRow As Long, lastCol As Long
    Dim n As Long, i As Long, j As Long, m As Long, pairs As Long
    Dim del() As Boolean, kinds() As Byte
    Dim dOT As Object, dOS As Object, dOR As Object, dOM As Object, groups As Object
    Dim typ As String, merch As String, key As String, amt As Double
    Dim k As Variant
    Dim lOT As Collection, lOS As Collection, lOR As Collection, lOM As Collection
    Dim charges As Collection, rets As Collection
    Dim cntOS As Long, cntRet As Long, skipped As Long, total As Long
    Dim msg As String

    On Error GoTo ErrHandler

    Set ws = DataSheet()
    firstRow = HEADER_ROW + 1
    lastRow = LastDataRow(ws)
    If lastRow < firstRow Then
        MsgBox "There is no data to check.", vbInformation, "Duplicates"
        Exit Sub
    End If
    lastCol = LastDataCol(ws)

    data = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).Value
    n = UBound(data, 1)
    ReDim del(1 To n)
    kinds = ScanMarks(ws, firstRow, lastRow)

    Set dOT = NewDict(): Set dOS = NewDict()
    Set dOR = NewDict(): Set dOM = NewDict()
    Set groups = NewDict()

    ' 1) Group rows by amount + merchant (rows already marked by Clean are ignored)
    For i = 1 To n
        If kinds(i) <> MARK_CLEAN Then
            typ = NormText(data(i, COL_TYPE))
            If typ = "OT" Or typ = "OS" Or typ = "OR" Or typ = "OM" Then
                merch = NormText(data(i, COL_MERCHANT))
                If TryAmount(data(i, COL_AMOUNT), amt) And merch <> "" Then
                    key = Format$(Abs(amt), "0.00") & "|" & merch
                    groups(key) = True
                    Select Case typ
                        Case "OT": AddRow dOT, key, i
                        Case "OS": AddRow dOS, key, i
                        Case "OR": AddRow dOR, key, i
                        Case "OM": AddRow dOM, key, i
                    End Select
                Else
                    skipped = skipped + 1
                End If
            End If
        End If
    Next i

    ' 2) Apply the rules inside each group
    For Each k In groups.Keys
        Set lOT = ListFor(dOT, CStr(k))
        Set lOS = ListFor(dOS, CStr(k))
        Set lOR = ListFor(dOR, CStr(k))
        Set lOM = ListFor(dOM, CStr(k))

        ' Rule A: each OT cancels one OS (OS marked, OT kept)
        pairs = lOT.Count
        If lOS.Count < pairs Then pairs = lOS.Count
        For j = 1 To pairs
            del(lOS(j)) = True
            cntOS = cntOS + 1
        Next j

        ' Rule B: each return (OR/OM) cancels one remaining OT/OS
        Set charges = New Collection
        If RETURN_CONSUMES_OT_FIRST Then
            AppendItems charges, lOT, 1
            AppendItems charges, lOS, pairs + 1
        Else
            AppendItems charges, lOS, pairs + 1
            AppendItems charges, lOT, 1
        End If

        Set rets = New Collection
        AppendItems rets, lOR, 1
        AppendItems rets, lOM, 1

        m = rets.Count
        If charges.Count < m Then m = charges.Count
        For j = 1 To m
            del(charges(j)) = True
            del(rets(j)) = True
        Next j
        cntRet = cntRet + m
    Next k

    total = cntOS + cntRet * 2
    ApplyMarks ws, del, MARK_DUP, firstRow, lastRow, lastCol    ' also clears stale Duplicate marks
    LogSecurityEvent gUsername, "Duplicates: highlighted " & total & " rows (OS/OT: " & cntOS & ", returns: " & cntRet & " pairs)"
    LoadList lv

    If total = 0 Then
        MsgBox "No duplicates found.", vbInformation, "Duplicates"
    Else
        msg = total & " rows highlighted (orange):" & vbCrLf & _
              "   OS matched with an OT:  " & cntOS & vbCrLf & _
              "   Returns matched with OT/OS:  " & cntRet & " pairs (" & cntRet * 2 & " rows)" & vbCrLf
        If skipped > 0 Then
            msg = msg & vbCrLf & skipped & " rows were ignored (invalid amount or blank merchant)." & vbCrLf
        End If
        msg = msg & vbCrLf & "Nothing was deleted. Review them, then press Delete."
        MsgBox msg, vbInformation, "Duplicates"
    End If
    Exit Sub

ErrHandler:
    MsgBox "Duplicates error: " & Err.Number & " - " & Err.Description, vbExclamation, "Duplicates"
End Sub

' ===================== DELETE HIGHLIGHTED ROWS =====================
Public Sub DeleteMarked(ByVal lv As MSComctlLib.ListView)
    Dim ws As Worksheet, data As Variant, kinds() As Byte
    Dim firstRow As Long, lastRow As Long, lastCol As Long
    Dim n As Long, i As Long
    Dim del() As Boolean, reason() As String
    Dim cClean As Long, cDup As Long, total As Long

    On Error GoTo ErrHandler

    Set ws = DataSheet()
    firstRow = HEADER_ROW + 1
    lastRow = LastDataRow(ws)
    If lastRow < firstRow Then
        MsgBox "There is no data.", vbInformation, "Delete"
        Exit Sub
    End If
    lastCol = LastDataCol(ws)

    kinds = ScanMarks(ws, firstRow, lastRow)
    data = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).Value
    n = UBound(data, 1)
    ReDim del(1 To n)
    ReDim reason(1 To n)

    For i = 1 To n
        Select Case kinds(i)
            Case MARK_CLEAN
                del(i) = True
                reason(i) = "Clean: highlighted and deleted by user"
                cClean = cClean + 1
            Case MARK_DUP
                del(i) = True
                reason(i) = "Duplicates: highlighted and deleted by user"
                cDup = cDup + 1
        End Select
    Next i

    total = cClean + cDup
    If total = 0 Then
        MsgBox "There are no highlighted rows to delete." & vbCrLf & _
               "Run Clean or Duplicates first.", vbInformation, "Delete"
        Exit Sub
    End If

    If MsgBox("Highlighted rows that will be deleted:" & vbCrLf & _
              "   From Clean (yellow):  " & cClean & vbCrLf & _
              "   From Duplicates (orange):  " & cDup & vbCrLf & _
              "   Total:  " & total & vbCrLf & vbCrLf & _
              "A copy is saved in '" & SHEET_REMOVED & "'. Continue?", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Delete") <> vbYes Then Exit Sub

    ArchiveAndDelete ws, data, del, reason, firstRow, lastCol
    LogSecurityEvent gUsername, "Delete: removed " & total & " highlighted rows (Clean: " & cClean & ", Duplicates: " & cDup & ")"
    LoadList lv
    MsgBox total & " rows deleted.", vbInformation, "Delete"
    Exit Sub

ErrHandler:
    MsgBox "Delete error: " & Err.Number & " - " & Err.Description, vbExclamation, "Delete"
End Sub
Public Sub LoadList(ByVal lv As MSComctlLib.ListView, Optional ByVal term As String = "")
    Dim ws As Worksheet, data As Variant, hdr As Variant, marks() As Byte
    Dim lastRow As Long, lastCol As Long, i As Long, j As Long
    Dim shown As Long, hit As Boolean, truncated As Boolean
    Dim itm As MSComctlLib.ListItem, txt As String

    Set ws = DataSheet()
    lastRow = LastDataRow(ws)
    lastCol = LastDataCol(ws)

    lv.View = lvwReport
    lv.FullRowSelect = True
    lv.Gridlines = True
    lv.ListItems.Clear
    lv.ColumnHeaders.Clear

    hdr = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastCol)).Value
    For j = 1 To lastCol
        txt = CellText(hdr(1, j))
        If txt = "" Then txt = "Col " & j
        lv.ColumnHeaders.Add , , txt, 90
    Next j

    If lastRow <= HEADER_ROW Then Exit Sub
    data = ws.Range(ws.Cells(HEADER_ROW + 1, 1), ws.Cells(lastRow, lastCol)).Value
    marks = ScanMarks(ws, HEADER_ROW + 1, lastRow)

    For i = 1 To UBound(data, 1)
        hit = (term = "")
        If Not hit Then
            For j = 1 To lastCol
                If InStr(1, CellText(data(i, j)), term, vbTextCompare) > 0 Then
                    hit = True
                    Exit For
                End If
            Next j
        End If

        If hit Then
            If shown >= MAX_LIST_ROWS Then
                truncated = True
                Exit For
            End If
            Set itm = lv.ListItems.Add(, , CellText(data(i, 1)))
            For j = 2 To lastCol
                itm.SubItems(j - 1) = CellText(data(i, j))
            Next j

            If marks(i) <> 0 Then            ' highlighted for deletion
                itm.ForeColor = RGB(192, 0, 0)
                itm.Bold = True
                For j = 2 To lastCol
                    itm.ListSubItems(j - 1).ForeColor = RGB(192, 0, 0)
                    itm.ListSubItems(j - 1).Bold = True
                Next j
            End If
            shown = shown + 1
        End If
    Next i

    If truncated Then
        MsgBox "Showing the first " & MAX_LIST_ROWS & " rows only. Refine the search to see others.", _
               vbInformation, "Workstation"
    End If
End Sub

' ===================== IMPORT (still a placeholder) =====================
' ===================== IMPORT =====================
Public Sub ImportData(ByVal lv As MSComctlLib.ListView)
    Application.Visible = False      ' Excel stays in the background
    ImportDataCore lv
    Application.Visible = False
End Sub
Private Function ReadExcelFile(ByVal path As String) As Variant
    Dim xl As Object, wb As Object, ws As Object
    Dim lr As Long, lc As Long
    Dim errNum As Long, errDesc As String

    On Error GoTo Fail
    Set xl = CreateObject("Excel.Application")   ' separate instance: your main Excel is never touched
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AutomationSecurity = 3                    ' no macros
    Set wb = xl.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set ws = wb.Worksheets(1)
    lr = LastDataRow(ws)
    lc = LastDataCol(ws)
    ReadExcelFile = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value

Cleanup:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close saveChanges:=False
    If Not xl Is Nothing Then xl.Quit
    Set ws = Nothing: Set wb = Nothing: Set xl = Nothing
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "modWorkstation", errDesc
    Exit Function

Fail:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Function
' ---------------- type helpers ----------------
Private Function InColList(ByVal col As Long, ByVal list As String) As Boolean
    InColList = (InStr("," & list & ",", "," & col & ",") > 0)
End Function

Private Function ColKind(ByVal col As Long) As String
    If InColList(col, NUMBER_COLS) Then
        ColKind = "N"
    ElseIf InColList(col, DATE_COLS) Then
        ColKind = "D"
    Else
        ColKind = "T"
    End If
End Function

Private Function CleanText(ByVal v As Variant) As String
    CleanText = Trim$(Replace(CStr(v), Chr$(160), " "))
End Function

' Returns the converted value. ok = False when it could not be converted
' (the original text is returned so nothing is lost).
Private Function ConvertCell(ByVal v As Variant, ByVal kind As String, ByRef ok As Boolean) As Variant
    Dim s As String, d As Double

    ok = True
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    Select Case kind
        Case "N"
            If VarType(v) = vbDate Then
                ConvertCell = CleanText(v): ok = False
            ElseIf VarType(v) <> vbString And IsNumeric(v) Then
                ConvertCell = CDbl(v)
            Else
                s = CleanText(v)
                If s = "" Then Exit Function
                If ParseNumber(s, d) Then
                    ConvertCell = d
                Else
                    ConvertCell = s: ok = False
                End If
            End If

        Case "D"
            If VarType(v) = vbDate Then
                ConvertCell = v
            ElseIf VarType(v) <> vbString And IsNumeric(v) Then
                If CDbl(v) >= 1 And CDbl(v) < 2958466 Then
                    ConvertCell = CDate(CDbl(v))
                Else
                    ConvertCell = CStr(v): ok = False
                End If
            Else
                s = CleanText(v)
                If s = "" Then Exit Function
                If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" And IsNumeric(Left$(s, 4)) _
                   And IsNumeric(Mid$(s, 6, 2)) And IsNumeric(Right$(s, 2)) Then
                    ConvertCell = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Right$(s, 2)))
                ElseIf IsDate(s) Then
                    ConvertCell = CDate(s)
                Else
                    ConvertCell = s: ok = False
                End If
            End If

        Case Else   ' text
            If VarType(v) = vbDouble Then
                If v = Int(v) And Abs(v) < 1E+15 Then
                    ConvertCell = Format$(v, "0")
                Else
                    ConvertCell = CStr(v)
                End If
            ElseIf VarType(v) = vbDate Then
                ConvertCell = Format$(v, "yyyy-mm-dd")
            Else
                ConvertCell = CleanText(v)
            End If
    End Select
End Function

' Handles $1,234.50 | 1.234,50 | (88.00) | 88.00- | -88
Private Function ParseNumber(ByVal s As String, ByRef d As Double) As Boolean
    Dim neg As Boolean, i As Long, nDot As Long, nCom As Long
    Dim norm As String, after As Long

    s = Replace(Replace(Replace(s, "$", ""), ChrW(8364), ""), " ", "")
    If Len(s) = 0 Then Exit Function

    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then neg = True: s = Mid$(s, 2, Len(s) - 2)
    If Right$(s, 1) = "-" Then neg = Not neg: s = Left$(s, Len(s) - 1)
    If Left$(s, 1) = "-" Then
        neg = Not neg: s = Mid$(s, 2)
    ElseIf Left$(s, 1) = "+" Then
        s = Mid$(s, 2)
    End If
    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "[0-9.,]") Then Exit Function
    Next i

    nDot = Len(s) - Len(Replace(s, ".", ""))
    nCom = Len(s) - Len(Replace(s, ",", ""))

    If nDot > 0 And nCom > 0 Then
        If InStrRev(s, ".") > InStrRev(s, ",") Then
            norm = Replace(s, ",", "")                              ' 1,234.50
        Else
            norm = Replace(Replace(s, ".", ""), ",", ".")           ' 1.234,50
        End If
    ElseIf nCom > 0 Then
        after = Len(s) - InStrRev(s, ",")
        If nCom > 1 Or after = 3 Then
            norm = Replace(s, ",", "")                              ' 1,234 = thousands
        Else
            norm = Replace(s, ",", ".")                             ' 88,5 = decimal
        End If
    ElseIf nDot > 1 Then
        norm = Replace(s, ".", "")
    Else
        norm = s
    End If

    If Not (norm Like "*[0-9]*") Then Exit Function
    d = Val(norm)
    If neg Then d = -d
    ParseNumber = True
End Function

' ---------------- file readers ----------------
Private Function ReadSource(ByVal path As String) As Variant
    Dim ext As String
    ext = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    Select Case ext
        Case "csv", "txt": ReadSource = ReadCsvFile(path)
        Case "xlsx", "xls", "xlsm": ReadSource = ReadExcelFile(path)
        Case Else
            Err.Raise vbObjectError + 514, "modWorkstation", "File type '" & ext & "' is not supported."
    End Select
End Function

Private Function ReadCsvFile(ByVal path As String) As Variant
    Dim stm As Object, txt As String
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                                 ' text
    stm.Charset = CSV_CHARSET
    stm.Open
    stm.LoadFromFile path
    txt = stm.ReadText
    stm.Close
    ReadCsvFile = ParseCsv(txt)
End Function

Private Function DetectDelimiter(ByVal line As String) As String
    Dim cands As Variant, i As Long, n As Long, best As Long
    cands = Array(",", ";", vbTab, "|")
    DetectDelimiter = ","
    For i = 0 To 3
        n = Len(line) - Len(Replace(line, cands(i), ""))
        If n > best Then best = n: DetectDelimiter = cands(i)
    Next i
End Function

Private Function SplitCsvLine(ByVal s As String, ByVal delim As String) As String()
    Dim res() As String, n As Long, i As Long, ch As String
    Dim inQ As Boolean, cur As String

    ReDim res(0 To 0)
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(s, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            inQ = True
        ElseIf ch = delim Then
            res(n) = cur
            n = n + 1
            ReDim Preserve res(0 To n)
            cur = ""
        Else
            cur = cur & ch
        End If
        i = i + 1
    Loop
    res(n) = cur
    SplitCsvLine = res
End Function

Private Function ParseCsv(ByVal txt As String) As Variant
    Dim lines() As String, rows As Collection, f() As String
    Dim i As Long, r As Long, c As Long, maxC As Long
    Dim cur As String, delim As String, arr() As Variant

    Set rows = New Collection
    txt = Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(txt, vbLf)
    delim = DetectDelimiter(lines(0))

    i = 0
    Do While i <= UBound(lines)
        cur = lines(i)
        ' a quoted field can contain line breaks: keep joining while quotes are open
        Do While ((Len(cur) - Len(Replace(cur, """", ""))) Mod 2 = 1) And i < UBound(lines)
            i = i + 1
            cur = cur & vbLf & lines(i)
        Loop
        i = i + 1

        If Len(Trim$(cur)) > 0 Then
            f = SplitCsvLine(cur, delim)
            rows.Add f
            If UBound(f) + 1 > maxC Then maxC = UBound(f) + 1
        End If
    Loop

    If rows.Count = 0 Then Exit Function

    ReDim arr(1 To rows.Count, 1 To maxC)
    For r = 1 To rows.Count
        f = rows(r)
        For c = 0 To UBound(f)
            arr(r, c + 1) = f(c)
        Next c
    Next r
    ParseCsv = arr
End Function
Private Sub ImportDataCore(ByVal lv As MSComctlLib.ListView)
    Dim ws As Worksheet, fileVar As Variant, path As String, fname As String
    Dim src As Variant, out() As Variant, fin() As Variant
    Dim nRows As Long, nCols As Long, firstSrc As Long
    Dim cnt As Long, r As Long, c As Long
    Dim kind As String, ok As Boolean, v As Variant, isBlank As Boolean
    Dim bad As Object, k As Variant, parts() As String
    Dim badNum As Long, badDate As Long, badType As Long, blanks As Long
    Dim firstRow As Long, lastRow As Long, lastCol As Long, startRow As Long, existing As Long
    Dim replaceMode As Boolean, hdrBlank As Boolean
    Dim data As Variant, del() As Boolean, reason() As String, i As Long
    Dim ans As VbMsgBoxResult, msg As String, rng As Range, shownBad As Long
    Dim prevCalc As XlCalculation, errNum As Long, errDesc As String

    On Error GoTo ErrHandler

    Set ws = DataSheet()

    fileVar = Application.GetOpenFilename( _
        "Data files (*.csv;*.txt;*.xlsx;*.xls;*.xlsm),*.csv;*.txt;*.xlsx;*.xls;*.xlsm", _
        , "Select the file to import")
    If VarType(fileVar) = vbBoolean Then Exit Sub
    path = CStr(fileVar)
    fname = Mid$(path, InStrRev(path, "\") + 1)

    src = ReadSource(path)
    If IsEmpty(src) Then
        MsgBox "The file has no data.", vbInformation, "Import"
        Exit Sub
    End If

    nRows = UBound(src, 1)
    nCols = UBound(src, 2)
    firstSrc = IIf(SOURCE_HAS_HEADER, 2, 1)

    If nRows < firstSrc Then
        MsgBox "The file has a header but no data rows.", vbInformation, "Import"
        Exit Sub
    End If
    If nCols < COL_CHECK Then
        MsgBox "The file has only " & nCols & " columns. At least " & COL_CHECK & _
               " are expected, in the same order as the Workstation sheet.", vbExclamation, "Import"
        Exit Sub
    End If

    ' ---- Convert every cell to its data type ----
    ReDim out(1 To nRows - firstSrc + 1, 1 To nCols)
    Set bad = NewDict()

    For r = firstSrc To nRows
        isBlank = True
        For c = 1 To nCols
            If Len(Trim$(CellText(src(r, c)))) > 0 Then isBlank = False: Exit For
        Next c

        If isBlank Then
            blanks = blanks + 1
        Else
            cnt = cnt + 1
            For c = 1 To nCols
                kind = ColKind(c)
                v = ConvertCell(src(r, c), kind, ok)

                If c = COL_TYPE And Not IsEmpty(v) Then
                    v = UCase$(CStr(v))
                    If v <> "" Then
                        If InStr("|OT|OS|OR|OM|", "|" & v & "|") = 0 Then badType = badType + 1
                    End If
                End If

                If Not ok Then
                    bad(cnt & "|" & c) = True
                    If kind = "N" Then badNum = badNum + 1 Else badDate = badDate + 1
                End If
                out(cnt, c) = v
            Next c
        End If
    Next r

    If cnt = 0 Then
        MsgBox "The file has only blank rows.", vbInformation, "Import"
        Exit Sub
    End If

    If blanks > 0 Then
        ReDim fin(1 To cnt, 1 To nCols)
        For r = 1 To cnt
            For c = 1 To nCols
                fin(r, c) = out(r, c)
            Next c
        Next r
        out = fin
    End If

    ' ---- Existing data? ----
    firstRow = HEADER_ROW + 1
    lastRow = LastDataRow(ws)
    existing = lastRow - HEADER_ROW
    If existing < 0 Then existing = 0
    hdrBlank = (Len(CellText(ws.Cells(HEADER_ROW, 1).Value)) = 0)

    msg = "File: " & fname & vbCrLf & _
          "Rows to import:  " & cnt & vbCrLf
    If blanks > 0 Then msg = msg & "Blank rows skipped:  " & blanks & vbCrLf
    If badNum + badDate + badType > 0 Then
        msg = msg & vbCrLf & "Warnings (rows are still imported):" & vbCrLf
        If badNum > 0 Then msg = msg & "   Numbers not recognized:  " & badNum & vbCrLf
        If badDate > 0 Then msg = msg & "   Dates not recognized:  " & badDate & vbCrLf
        If badType > 0 Then msg = msg & "   Type (F) not OT/OS/OR/OM:  " & badType & vbCrLf
        If badNum + badDate > 0 Then msg = msg & "   (red cells need review)" & vbCrLf
    End If

    If existing > 0 Then
        msg = msg & vbCrLf & "The sheet already has " & existing & " rows." & vbCrLf & _
              "YES = add to them" & vbCrLf & _
              "NO = replace them (old rows are saved in '" & SHEET_REMOVED & "')" & vbCrLf & _
              "CANCEL = stop"
        ans = MsgBox(msg, vbYesNoCancel + vbQuestion, "Import")
        If ans = vbCancel Then Exit Sub
        replaceMode = (ans = vbNo)
    Else
        msg = msg & vbCrLf & "Continue?"
        If MsgBox(msg, vbYesNo + vbQuestion, "Import") <> vbYes Then Exit Sub
    End If

    ' ---- Replace: archive old rows first ----
    If replaceMode Then
        lastCol = LastDataCol(ws)
        data = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).Value
        ReDim del(1 To existing)
        ReDim reason(1 To existing)
        For i = 1 To existing
            del(i) = True
            reason(i) = "Import: replaced by " & fname
        Next i
        ArchiveAndDelete ws, data, del, reason, firstRow, lastCol
        startRow = firstRow
    ElseIf existing > 0 Then
        startRow = lastRow + 1
    Else
        startRow = firstRow
    End If

    ' ---- Write everything at once ----
    prevCalc = Application.Calculation
    On Error GoTo WriteFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If hdrBlank And SOURCE_HAS_HEADER Then
        For c = 1 To nCols
            ws.Cells(HEADER_ROW, c).Value = CellText(src(1, c))
        Next c
        ws.rows(HEADER_ROW).Font.Bold = True
    End If

    ' Set the format BEFORE writing so text stays text
    For c = 1 To nCols
        Set rng = ws.Range(ws.Cells(startRow, c), ws.Cells(startRow + cnt - 1, c))
        Select Case ColKind(c)
            Case "N": rng.NumberFormat = "#,##0.00"
            Case "D": rng.NumberFormat = "yyyy-mm-dd"
            Case Else: rng.NumberFormat = "@"
        End Select
    Next c

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + cnt - 1, nCols)).Value = out

    For Each k In bad.Keys
        If shownBad >= 5000 Then Exit For
        parts = Split(CStr(k), "|")
        ws.Cells(startRow + CLng(parts(0)) - 1, CLng(parts(1))).Interior.Color = RGB(255, 199, 206)
        shownBad = shownBad + 1
    Next k

WriteCleanup:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    On Error GoTo 0                      ' so the Raise below is not caught again
    If errNum <> 0 Then Err.Raise errNum, "modWorkstation", errDesc

    On Error GoTo ErrHandler
    LogSecurityEvent gUsername, "Import: " & cnt & " rows from " & fname & _
                     IIf(replaceMode, " (replace)", " (add)") & _
                     ", warnings: " & (badNum + badDate + badType)
    LoadList lv
    MsgBox cnt & " rows imported.", vbInformation, "Import"
    Exit Sub

WriteFail:
    errNum = Err.Number
    errDesc = Err.Description
    Resume WriteCleanup

ErrHandler:
    LogSecurityEvent gUsername, "Import FAILED: " & Err.Description
    MsgBox "Import error: " & Err.Number & " - " & Err.Description, vbExclamation, "Import"
End Sub
' ===================== ARCHIVE + DELETE =====================
Private Sub ArchiveAndDelete(ByVal ws As Worksheet, ByRef data As Variant, _
                             ByRef del() As Boolean, ByRef reason() As String, _
                             ByVal firstRow As Long, ByVal lastCol As Long)
    Dim wsLog As Worksheet, out() As Variant
    Dim n As Long, i As Long, j As Long, k As Long, startRow As Long
    Dim prevCalc As XlCalculation
    Dim errNum As Long, errDesc As String

    n = UBound(del)
    For i = 1 To n
        If del(i) Then k = k + 1
    Next i
    If k = 0 Then Exit Sub

    ReDim out(1 To k, 1 To lastCol + 4)
    k = 0
    For i = 1 To n
        If del(i) Then
            k = k + 1
            out(k, 1) = Now
            out(k, 2) = gUsername
            out(k, 3) = reason(i)
            out(k, 4) = firstRow + i - 1        ' original row number
            For j = 1 To lastCol
                out(k, 4 + j) = data(i, j)
            Next j
        End If
    Next i

    prevCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Save a copy first, then delete
    Set wsLog = RemovedSheet(ws, lastCol)
    startRow = wsLog.Cells(wsLog.rows.Count, 1).End(xlUp).Row + 1
    wsLog.Range(wsLog.Cells(startRow, 1), wsLog.Cells(startRow + k - 1, 1)).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wsLog.Range(wsLog.Cells(startRow, 1), wsLog.Cells(startRow + k - 1, lastCol + 4)).Value = out

    DeleteMarkedRows ws, del, firstRow, lastCol

Cleanup:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "modWorkstation", errDesc
    Exit Sub

Fail:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Sub

' Fast delete: flag column + AutoFilter + delete visible rows
Private Sub DeleteMarkedRows(ByVal ws As Worksheet, ByRef del() As Boolean, _
                             ByVal firstRow As Long, ByVal lastCol As Long)
    Dim n As Long, i As Long, flagCol As Long, lastRow As Long
    Dim flags() As Variant
    Dim errNum As Long, errDesc As String

    n = UBound(del)
    lastRow = firstRow + n - 1
    flagCol = lastCol + 1

    ReDim flags(1 To n, 1 To 1)
    For i = 1 To n
        If del(i) Then flags(i, 1) = "DEL"
    Next i

    On Error GoTo Fail
    ws.AutoFilterMode = False
    ws.Cells(HEADER_ROW, flagCol).Value = "_flag"
    ws.Range(ws.Cells(firstRow, flagCol), ws.Cells(lastRow, flagCol)).Value = flags
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, flagCol)).AutoFilter Field:=flagCol, Criteria1:="DEL"
    ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, flagCol)).SpecialCells(xlCellTypeVisible).EntireRow.Delete

Cleanup:
    On Error Resume Next
    ws.AutoFilterMode = False
    ws.Columns(flagCol).Clear
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, "modWorkstation", errDesc
    Exit Sub

Fail:
    errNum = Err.Number
    errDesc = Err.Description
    Resume Cleanup
End Sub

Private Function RemovedSheet(ByVal src As Worksheet, ByVal lastCol As Long) As Worksheet
    Dim ws As Worksheet, prev As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_REMOVED)
    On Error GoTo 0

    If ws Is Nothing Then
        Set prev = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_REMOVED
        ws.Range("A1:D1").Value = Array("Removed on", "Removed by", "Reason", "Source row")
        ws.Cells(1, 5).Resize(1, lastCol).Value = _
            src.Range(src.Cells(HEADER_ROW, 1), src.Cells(HEADER_ROW, lastCol)).Value
        ws.rows(1).Font.Bold = True
        On Error Resume Next
        prev.Activate
        On Error GoTo 0
    End If
    Set RemovedSheet = ws
End Function
' Removes all data rows from the sheet (copied to RemovedRows first)
Public Sub ClearSheetData(ByVal lv As MSComctlLib.ListView)
    Dim ws As Worksheet, data As Variant
    Dim firstRow As Long, lastRow As Long, lastCol As Long
    Dim n As Long, i As Long
    Dim del() As Boolean, reason() As String

    On Error GoTo ErrHandler

    Set ws = DataSheet()
    firstRow = HEADER_ROW + 1
    lastRow = LastDataRow(ws)

    If lastRow < firstRow Then
        lv.ListItems.Clear
        Exit Sub
    End If

    n = lastRow - firstRow + 1
    If MsgBox("This will remove all " & n & " rows from the Workstation sheet." & vbCrLf & _
              "A copy is saved in '" & SHEET_REMOVED & "'. Continue?", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Clear") <> vbYes Then Exit Sub

    lastCol = LastDataCol(ws)
    data = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).Value

    ReDim del(1 To n)
    ReDim reason(1 To n)
    For i = 1 To n
        del(i) = True
        reason(i) = "Clear: data cleared by user"
    Next i

    ArchiveAndDelete ws, data, del, reason, firstRow, lastCol
    LogSecurityEvent gUsername, "Clear: removed " & n & " rows from Workstation"
    lv.ListItems.Clear
    Exit Sub

ErrHandler:
    MsgBox "Clear error: " & Err.Number & " - " & Err.Description, vbExclamation, "Clear"
End Sub

