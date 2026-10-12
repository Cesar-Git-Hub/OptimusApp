VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "Optimus Dashboard"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10590
   OleObjectBlob   =   "frmDashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLogout_Click()
    LogSecurityEvent gUsername, "Logout"
    gUsername = ""
    gUserRole = ""
    Unload Me
    frmLogin.Show
End Sub

Private Sub UserForm_Initialize()
    ' Admin-only features are disabled for everyone else
    cmdAdmin.Enabled = IsAdmin()
    cmdSecurity.Enabled = IsAdmin()

    ' Optional: add a Label named lblWelcome to the form, then uncomment
    ' lblWelcome.Caption = "Welcome, " & gUsername & " (" & gUserRole & ")"
End Sub

' ---------------------------------------------------------
' Open a child form: hide the Dashboard, show it again afterwards
' ---------------------------------------------------------
Private Sub cmdWorkstation_Click()
    On Error GoTo ErrHandler
    LogSecurityEvent gUsername, "Opened Workstation"
    Me.Hide
    frmWorkstation.Show
    If Not gExiting Then Me.Show
    Exit Sub
ErrHandler:
    Me.Show
    MsgBox "Error opening Workstation: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Workstation"
End Sub

Private Sub cmdAdmin_Click()
    On Error GoTo ErrHandler

    ' Re-check the role even though the button is disabled
    If Not IsAdmin() Then
        LogSecurityEvent gUsername, "Denied: Admin panel"
        MsgBox "This function requires administrator permissions.", _
               vbExclamation, "Access denied"
        Exit Sub
    End If

    LogSecurityEvent gUsername, "Opened Admin panel"
    Me.Hide
    frmAdmin.Show
    If Not gExiting Then Me.Show
    Exit Sub
ErrHandler:
    Me.Show
    MsgBox "Error opening Admin panel: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Admin"
End Sub

Private Sub cmdSecurity_Click()
    On Error GoTo ErrHandler

    If Not IsAdmin() Then
        LogSecurityEvent gUsername, "Denied: Security module"
        MsgBox "This function requires administrator permissions.", _
               vbExclamation, "Access denied"
        Exit Sub
    End If

    LogSecurityEvent gUsername, "Opened Security module"
    Me.Hide
    frmSecurity.Show
    If Not gExiting Then Me.Show
    Exit Sub

ErrHandler:
    Me.Show
    MsgBox "Error opening Security module: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Security"
End Sub
Private Sub cmdExit_Click()
    If MsgBox("Do you want to exit Optimus Fraud?", vbQuestion + vbYesNo, "Exit") = vbYes Then
        ExitApplication
    End If
End Sub
