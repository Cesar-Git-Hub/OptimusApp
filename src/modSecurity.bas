Attribute VB_Name = "modSecurity"
' ============ SHA-256 in pure VBA (no .NET needed) ============

Public Function HashPassword(ByVal username As String, ByVal plainText As String) As String
    HashPassword = SHA256Hex(LCase$(Trim$(username)) & "|" & plainText)
End Function

Private Function AddU(ByVal a As Long, ByVal b As Long) As Long
    Dim lo As Long, hi As Long
    lo = (a And &HFFFF&) + (b And &HFFFF&)
    hi = (((a And &HFFFF0000) \ &H10000) And &HFFFF&) + _
         (((b And &HFFFF0000) \ &H10000) And &HFFFF&) + (lo \ &H10000)
    hi = hi And &HFFFF&
    lo = lo And &HFFFF&
    If hi And &H8000& Then
        AddU = ((hi And &H7FFF&) * &H10000) Or lo Or &H80000000
    Else
        AddU = (hi * &H10000) Or lo
    End If
End Function

Private Function ShR(ByVal x As Long, ByVal n As Long) As Long
    If n = 0 Then
        ShR = x
    ElseIf x >= 0 Then
        ShR = x \ (2 ^ n)
    Else
        ShR = ((x And &H7FFFFFFF) \ (2 ^ n)) Or (2 ^ (31 - n))
    End If
End Function

Private Function ShL(ByVal x As Long, ByVal n As Long) As Long
    Dim r As Long
    r = (x And ((2 ^ (31 - n)) - 1)) * (2 ^ n)
    If (x And (2 ^ (31 - n))) <> 0 Then r = r Or &H80000000
    ShL = r
End Function

Private Function RotR(ByVal x As Long, ByVal n As Long) As Long
    RotR = ShR(x, n) Or ShL(x, 32 - n)
End Function

Private Function Utf8Bytes(ByVal s As String) As Byte()
    Dim b() As Byte, i As Long, n As Long, c As Long
    ReDim b(0 To Len(s) * 3 + 3)
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        If c < &H80 Then
            b(n) = c: n = n + 1
        ElseIf c < &H800 Then
            b(n) = &HC0 Or (c \ 64): b(n + 1) = &H80 Or (c And 63): n = n + 2
        Else
            b(n) = &HE0 Or (c \ 4096)
            b(n + 1) = &H80 Or ((c \ 64) And 63)
            b(n + 2) = &H80 Or (c And 63)
            n = n + 3
        End If
    Next i
    ReDim Preserve b(0 To n - 1)
    Utf8Bytes = b
End Function

Private Function SHA256Hex(ByVal s As String) As String
    Dim k As Variant, H(0 To 7) As Long, w(0 To 63) As Long
    Dim msg() As Byte, n As Long, total As Long, bits As Long
    Dim blk As Long, t As Long, j As Long, base As Long
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, hh As Long
    Dim s0 As Long, s1 As Long, ch As Long, maj As Long, t1 As Long, t2 As Long
    Dim result As String

    k = Array( _
        &H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)

    H(0) = &H6A09E667: H(1) = &HBB67AE85: H(2) = &H3C6EF372: H(3) = &HA54FF53A
    H(4) = &H510E527F: H(5) = &H9B05688C: H(6) = &H1F83D9AB: H(7) = &H5BE0CD19

    msg = Utf8Bytes(s)
    n = UBound(msg) + 1
    total = ((n + 8) \ 64 + 1) * 64
    ReDim Preserve msg(0 To total - 1)
    msg(n) = &H80
    bits = n * 8
    msg(total - 4) = (bits \ &H1000000) And &HFF
    msg(total - 3) = (bits \ &H10000) And &HFF
    msg(total - 2) = (bits \ &H100) And &HFF
    msg(total - 1) = bits And &HFF

    For blk = 0 To (total \ 64) - 1
        base = blk * 64
        For t = 0 To 15
            j = base + t * 4
            w(t) = ((msg(j) And &H7F) * &H1000000) Or (msg(j + 1) * &H10000) Or _
                   (msg(j + 2) * &H100&) Or msg(j + 3)
            If msg(j) And &H80 Then w(t) = w(t) Or &H80000000
        Next t
        For t = 16 To 63
            s0 = RotR(w(t - 15), 7) Xor RotR(w(t - 15), 18) Xor ShR(w(t - 15), 3)
            s1 = RotR(w(t - 2), 17) Xor RotR(w(t - 2), 19) Xor ShR(w(t - 2), 10)
            w(t) = AddU(AddU(w(t - 16), s0), AddU(w(t - 7), s1))
        Next t

        a = H(0): b = H(1): c = H(2): d = H(3)
        e = H(4): f = H(5): g = H(6): hh = H(7)

        For t = 0 To 63
            s1 = RotR(e, 6) Xor RotR(e, 11) Xor RotR(e, 25)
            ch = (e And f) Xor ((Not e) And g)
            t1 = AddU(AddU(AddU(hh, s1), AddU(ch, k(t))), w(t))
            s0 = RotR(a, 2) Xor RotR(a, 13) Xor RotR(a, 22)
            maj = (a And b) Xor (a And c) Xor (b And c)
            t2 = AddU(s0, maj)
            hh = g: g = f: f = e
            e = AddU(d, t1)
            d = c: c = b: b = a
            a = AddU(t1, t2)
        Next t

        H(0) = AddU(H(0), a): H(1) = AddU(H(1), b)
        H(2) = AddU(H(2), c): H(3) = AddU(H(3), d)
        H(4) = AddU(H(4), e): H(5) = AddU(H(5), f)
        H(6) = AddU(H(6), g): H(7) = AddU(H(7), hh)
    Next blk

    For t = 0 To 7
        result = result & Right$("00000000" & Hex$(H(t)), 8)
    Next t
    SHA256Hex = LCase$(result)
End Function

Public Sub LogSecurityEvent(ByVal username As String, ByVal action As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    On Error GoTo Failed

    Set ws = ThisWorkbook.Worksheets("UserSecurityLog")
    nextRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = username
    ws.Cells(nextRow, 3).Value = action
    Exit Sub

Failed:
    Debug.Print "LogSecurityEvent failed: " & Err.Number & " - " & Err.Description
End Sub
Public Function IsAdmin() As Boolean
    IsAdmin = (StrComp(gUserRole, "Admin", vbTextCompare) = 0)
End Function
