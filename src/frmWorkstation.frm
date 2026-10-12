VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWorkstation 
   Caption         =   "Optimus Workstation"
   ClientHeight    =   8835.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16035
   OleObjectBlob   =   "frmWorkstation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWorkstation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler
    LogSecurityEvent gUsername, "Workstation delete highlighted rows"
    modWorkstation.DeleteMarked Me.lstTransactions
    Exit Sub
ErrHandler:
    MsgBox "Delete error: " & Err.Number & " - " & Err.Description, vbExclamation, "Delete"
End Sub

' ---------------------------------------------------------
' Search
' ---------------------------------------------------------
Private Sub cmdSearch_Click()
    Dim term As String

    On Error GoTo ErrHandler

    term = Trim$(txtSearch.Value)
    If term = "" Then
        MsgBox "Enter a value to search.", vbExclamation, "Search"
        txtSearch.SetFocus
        Exit Sub
    End If

    LogSecurityEvent gUsername, "Workstation search: " & term
    modWorkstation.SearchData term, Me.lstTransactions
    Exit Sub

ErrHandler:
    MsgBox "Search error: " & Err.Number & " - " & Err.Description, vbExclamation, "Search"
End Sub

Private Sub cmdClear_Click()
    On Error GoTo ErrHandler
    modWorkstation.ClearSheetData Me.lstTransactions
    txtSearch.Value = ""
    txtSearch.SetFocus
    Exit Sub
ErrHandler:
    MsgBox "Clear error: " & Err.Number & " - " & Err.Description, vbExclamation, "Clear"
End Sub

' ---------------------------------------------------------
' Import / Clean / Duplicates
' ---------------------------------------------------------
Private Sub cmdImport_Click()
    On Error GoTo ErrHandler
    LogSecurityEvent gUsername, "Workstation import started"
    modWorkstation.ImportData Me.lstTransactions
    LogSecurityEvent gUsername, "Workstation import finished"
    Exit Sub
ErrHandler:
    LogSecurityEvent gUsername, "Workstation import FAILED: " & Err.Description
    MsgBox "Import error: " & Err.Number & " - " & Err.Description, vbExclamation, "Import"
End Sub

Private Sub cmdClean_Click()
    On Error GoTo ErrHandler
    modWorkstation.CleanData Me.lstTransactions
    Exit Sub
ErrHandler:
    MsgBox "Clean error: " & Err.Number & " - " & Err.Description, vbExclamation, "Clean"
End Sub

Private Sub cmdDuplicates_Click()
    On Error GoTo ErrHandler
    LogSecurityEvent gUsername, "Workstation duplicate check"
    modWorkstation.FindDuplicates Me.lstTransactions
    Exit Sub
ErrHandler:
    MsgBox "Duplicate check error: " & Err.Number & " - " & Err.Description, vbExclamation, "Duplicates"
End Sub

' ---------------------------------------------------------
' Navigation
' ---------------------------------------------------------
' The Dashboard is already waiting behind this form and
' reappears by itself when this form closes.
Private Sub cmdDashboard_Click()
    Unload Me
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Do you want to exit Optimus Fraud?", vbQuestion + vbYesNo, "Exit") = vbYes Then
        ExitApplication
    End If
End Sub

' Show the Workstation sheet in Excel (form hides while you work on it)
Private Sub cmdExcel_Click()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Workstation")
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        MsgBox "The 'Workstation' sheet doesn't exist in this workbook.", vbExclamation, "Excel"
        Exit Sub
    End If

    LogSecurityEvent gUsername, "Opened Workstation sheet"
    Me.Hide
    Application.Visible = True
    ws.Visible = xlSheetVisible
    ws.Activate
    Exit Sub

ErrHandler:
    Me.Show
    MsgBox "Error opening the sheet: " & Err.Number & " - " & Err.Description, vbExclamation, "Excel"
End Sub

