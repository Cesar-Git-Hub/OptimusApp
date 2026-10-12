VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSecurity 
   Caption         =   "Optimus Security"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
   OleObjectBlob   =   "frmSecurity.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SHEET_LOG As String = "UserSecurityLog"
Private Const MAX_ROWS As Long = 2000   ' newest entries first

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    lstLog.ColumnCount = 3
    lstLog.ColumnWidths = "120;90;300"
    LoadLog
    Exit Sub
ErrHandler:
    MsgBox "Error initializing Security: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Security"
End Sub

Private Sub LoadLog()
    Dim ws As Worksheet, data As Variant, arr() As Variant
    Dim lastRow As Long, i As Long, k As Long, n As Long

    lstLog.Clear

    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 3)).Value
    n = UBound(data, 1)
    If n > MAX_ROWS Then n = MAX_ROWS

    ReDim arr(0 To n - 1, 0 To 2)
    k = 0
    For i = UBound(data, 1) To UBound(data, 1) - n + 1 Step -1   ' newest first
        If IsDate(data(i, 1)) Then
            arr(k, 0) = Format$(data(i, 1), "yyyy-mm-dd hh:nn:ss")
        Else
            arr(k, 0) = CStr(data(i, 1))
        End If
        arr(k, 1) = CStr(data(i, 2))
        arr(k, 2) = CStr(data(i, 3))
        k = k + 1
    Next i

    lstLog.list = arr
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrHandler
    LoadLog
    Exit Sub
ErrHandler:
    MsgBox "Error refreshing: " & Err.Number & " - " & Err.Description, vbExclamation, "Security"
End Sub

Private Sub cmdDashboard_Click()
    Unload Me
End Sub
