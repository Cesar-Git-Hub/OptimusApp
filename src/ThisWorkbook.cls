VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Set to True while developing so the system sheets stay visible
Private Const DEV_MODE As Boolean = True

Private Sub Workbook_Open()
    Dim ws As Worksheet

    On Error GoTo ErrHandler

    If ThisWorkbook.ProtectStructure Then
        On Error Resume Next
        ThisWorkbook.Unprotect password:=""
        On Error GoTo ErrHandler
        If ThisWorkbook.ProtectStructure Then
            Debug.Print "Workbook structure is password protected; visibility not changed."
        End If
    End If

    ' Make sure the required sheets exist
    EnsureSheet "Blank"
    EnsureSheet "Users", Array("Username", "Password", "Role", "Active", "Last Login")
    EnsureSheet "UserSecurityLog", Array("Date", "User", "Action")

    ' Adjust visibility (Blank is always left visible)
        If Not ThisWorkbook.ProtectStructure Then
        For Each ws In ThisWorkbook.Worksheets
            Select Case ws.Name
                Case "Blank", "Workstation"
                    ws.Visible = xlSheetVisible
                Case "Users", "UserSecurityLog", "Dashboard"
                    If DEV_MODE Then
                        ws.Visible = xlSheetVisible
                    Else
                        ws.Visible = xlSheetVeryHidden
                    End If
            End Select
        Next ws
    End If
    If Not DEV_MODE And Workbooks.Count = 1 Then Application.Visible = False
    Application.OnTime Now + TimeSerial(0, 0, 1), "ShowLoginForm"
    Exit Sub

ErrHandler:
    MsgBox "Error in Workbook_Open: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Startup"
End Sub

' Creates the sheet (with headers) if it doesn't exist
Private Sub EnsureSheet(ByVal sheetName As String, Optional ByVal headers As Variant)
    Dim ws As Worksheet, i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        If Not IsMissing(headers) Then
            For i = LBound(headers) To UBound(headers)
                ws.Cells(1, i - LBound(headers) + 1).Value = headers(i)
            Next i
            ws.rows(1).Font.Bold = True
        End If
    End If
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If gExiting Then ThisWorkbook.Saved = True
End Sub
