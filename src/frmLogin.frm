VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Login form"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4740
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SHEET_USERS As String = "Users"
Private Const COL_USER As Long = 1
Private Const COL_PASS As Long = 2
Private Const COL_ROLE As Long = 3
Private Const COL_ACTIVE As Long = 4
Private Const COL_LASTLOGIN As Long = 5
Private Const MAX_ATTEMPTS As Long = 3

Private mAttempts As Long

Private Sub UserForm_Initialize()
    txtPassword.PasswordChar = "*"
    cmdLogin.Default = True        ' Enter key triggers Login
End Sub

' Accepts TRUE, 1, Yes, Si (with or without accent)
Private Function IsActiveValue(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    IsActiveValue = (s = "TRUE" Or s = "1" Or s = "YES" Or s = "SI" Or s = "S" & ChrW(205))
End Function

Private Sub cmdLogin_Click()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim u As String, p As String
    Dim userRow As Long

    On Error GoTo ErrHandler

    u = Trim$(txtUsername.Value)
    p = txtPassword.Value          ' do not trim passwords

    If u = "" Or p = "" Then
        MsgBox "You must enter a username and password.", vbExclamation, "Login"
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_USERS)
    lastRow = ws.Cells(ws.rows.Count, COL_USER).End(xlUp).Row

    ' Find the user
    For i = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(i, COL_USER).Value)), u, vbTextCompare) = 0 Then
            userRow = i
            Exit For
        End If
    Next i

    ' Validate: user exists, is active, password hash matches
    If userRow > 0 Then
        If IsActiveValue(ws.Cells(userRow, COL_ACTIVE).Value) Then
            If StrComp(CStr(ws.Cells(userRow, COL_PASS).Value), _
                       HashPassword(CStr(ws.Cells(userRow, COL_USER).Value), p), _
                       vbTextCompare) = 0 Then

                gUsername = Trim$(CStr(ws.Cells(userRow, COL_USER).Value))
                LogSecurityEvent gUsername, "Login Success"
                gUserRole = CStr(ws.Cells(userRow, COL_ROLE).Value)
                ws.Cells(userRow, COL_LASTLOGIN).Value = Now
                ThisWorkbook.Save

                Unload Me
                frmDashboard.Show
                Exit Sub
            End If
        End If
    End If

    ' Same message for every failure (no hints to attackers)
    mAttempts = mAttempts + 1
    LogSecurityEvent u, "Login Failed (attempt " & mAttempts & ")"
    txtPassword.Value = ""

    If mAttempts >= MAX_ATTEMPTS Then
        MsgBox "Too many failed attempts. The application will close.", vbCritical, "Login"
        CloseApplication
    Else
        MsgBox "Invalid username or password, or the account is inactive." & vbCrLf & _
               "Attempts left: " & (MAX_ATTEMPTS - mAttempts), vbCritical, "Login"
        txtPassword.SetFocus
    End If
    Exit Sub

ErrHandler:
    MsgBox "Login error: " & Err.Number & " - " & Err.Description, vbCritical, "Login"
End Sub

Private Sub cmdExit_Click()
    If MsgBox("Do you want to exit Optimus Fraud?", vbQuestion + vbYesNo, "Exit") = vbYes Then
        ExitApplication
    End If
End Sub

Private Sub CloseApplication()
    Unload Me
    If Workbooks.Count = 1 Then
        Application.Quit               ' only workbook open: safe to quit Excel
    Else
        ThisWorkbook.Close saveChanges:=True   ' don't kill the user's other workbooks
    End If
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        ExitApplication
    End If
End Sub
