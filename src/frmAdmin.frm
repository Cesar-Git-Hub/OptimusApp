VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdmin 
   Caption         =   "Administration"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10470
   OleObjectBlob   =   "frmAdmin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdmin"
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

' ---------------------------------------------------------
' Navigation
' ---------------------------------------------------------
Private Sub cmdDashboard_Click()
    On Error GoTo ErrHandler
    Unload Me
    frmDashboard.Show vbModal
    Exit Sub
ErrHandler:
    MsgBox "Error opening Dashboard: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "Dashboard"
End Sub

' ---------------------------------------------------------
' Initialization
' ---------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler

    With lstUsers
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Username", 110
        .ColumnHeaders.Add , , "Role", 80
        .ColumnHeaders.Add , , "Active", 60
        .ColumnHeaders.Add , , "Last Login", 130
    End With

    With cmbRole
        .Style = fmStyleDropDownList   ' only allow listed values
        .Clear
        .AddItem "Admin"
        .AddItem "User"
    End With

    txtPassword.PasswordChar = "*"

    LoadUsers
    ClearFields
    Exit Sub

ErrHandler:
    Debug.Print "Error in frmAdmin.Initialize - " & Err.Number & " - " & Err.Description
    MsgBox "An error occurred while initializing the Admin panel.", vbExclamation, "Admin"
End Sub

' ---------------------------------------------------------
' Helpers
' ---------------------------------------------------------
Private Function UsersSheet() As Worksheet
    Set UsersSheet = ThisWorkbook.Worksheets(SHEET_USERS)
End Function

' Returns the sheet row of a username, or 0 if not found
Private Function FindUserRow(ByVal username As String) As Long
    Dim ws As Worksheet, lastRow As Long, i As Long

    Set ws = UsersSheet()
    lastRow = ws.Cells(ws.rows.Count, COL_USER).End(xlUp).Row

    For i = 2 To lastRow
        If StrComp(CStr(ws.Cells(i, COL_USER).Value), username, vbTextCompare) = 0 Then
            FindUserRow = i
            Exit Function
        End If
    Next i
End Function

' Counts active admins (used to protect the last one)
Private Function CountActiveAdmins() As Long
    Dim ws As Worksheet, lastRow As Long, i As Long, n As Long

    Set ws = UsersSheet()
    lastRow = ws.Cells(ws.rows.Count, COL_USER).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, COL_ROLE).Value = "Admin" And ws.Cells(i, COL_ACTIVE).Value = "Yes" Then
            n = n + 1
        End If
    Next i
    CountActiveAdmins = n
End Function

Private Sub ClearFields()
    txtUsername.Value = ""
    txtPassword.Value = ""
    cmbRole.ListIndex = -1
    chkActive.Value = True
    txtUsername.Enabled = True
End Sub

' ---------------------------------------------------------
' Load users into the ListView (password is never shown)
' ---------------------------------------------------------
Private Sub LoadUsers()
    Dim ws As Worksheet, data As Variant
    Dim lastRow As Long, i As Long
    Dim itm As ListItem

    Set ws = UsersSheet()
    lstUsers.ListItems.Clear

    lastRow = ws.Cells(ws.rows.Count, COL_USER).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, COL_LASTLOGIN)).Value

    For i = 1 To UBound(data, 1)
        Set itm = lstUsers.ListItems.Add(, , CStr(data(i, COL_USER)))
        itm.SubItems(1) = CStr(data(i, COL_ROLE))
        itm.SubItems(2) = CStr(data(i, COL_ACTIVE))
        If IsDate(data(i, COL_LASTLOGIN)) Then
            itm.SubItems(3) = Format$(data(i, COL_LASTLOGIN), "yyyy-mm-dd hh:nn")
        Else
            itm.SubItems(3) = CStr(data(i, COL_LASTLOGIN))
        End If
    Next i
End Sub

' Fill the fields when a user is selected
Private Sub lstUsers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtUsername.Value = Item.Text
    txtUsername.Enabled = False            ' username can't be changed
    cmbRole.Value = Item.SubItems(1)
    chkActive.Value = (Item.SubItems(2) = "Yes")
    txtPassword.Value = ""                 ' blank = keep current password
End Sub

' ---------------------------------------------------------
' Add
' ---------------------------------------------------------
Private Sub cmdAdd_Click()
    Dim ws As Worksheet, newRow As Long
    Dim uname As String

    On Error GoTo ErrHandler

    uname = Trim$(txtUsername.Value)

    If uname = "" Or txtPassword.Value = "" Or cmbRole.ListIndex = -1 Then
        MsgBox "Please complete all fields.", vbExclamation, "Add User"
        Exit Sub
    End If

    If FindUserRow(uname) > 0 Then
        MsgBox "That username already exists.", vbExclamation, "Add User"
        Exit Sub
    End If

    Set ws = UsersSheet()
    newRow = ws.Cells(ws.rows.Count, COL_USER).End(xlUp).Row + 1

    ws.Cells(newRow, COL_USER).Value = uname
    ws.Cells(newRow, COL_PASS).Value = HashPassword(uname, txtPassword.Value)
    ws.Cells(newRow, COL_ROLE).Value = cmbRole.Value
    ws.Cells(newRow, COL_ACTIVE).Value = IIf(chkActive.Value, "Yes", "No")
    ws.Cells(newRow, COL_LASTLOGIN).Value = ""

    LoadUsers
    ClearFields
    MsgBox "User added.", vbInformation, "Add User"
    Exit Sub

ErrHandler:
    MsgBox "Error adding user: " & Err.Number & " - " & Err.Description, vbExclamation, "Add User"
End Sub

' ---------------------------------------------------------
' Update
' ---------------------------------------------------------
Private Sub cmdUpdate_Click()
    Dim ws As Worksheet, r As Long
    Dim uname As String, willBeActiveAdmin As Boolean

    On Error GoTo ErrHandler

    If lstUsers.SelectedItem Is Nothing Then
        MsgBox "Please select a user.", vbExclamation, "Update User"
        Exit Sub
    End If

    If cmbRole.ListIndex = -1 Then
        MsgBox "Please choose a role.", vbExclamation, "Update User"
        Exit Sub
    End If

    uname = lstUsers.SelectedItem.Text
    r = FindUserRow(uname)
    If r = 0 Then
        MsgBox "User not found. The list will be refreshed.", vbExclamation, "Update User"
        LoadUsers
        Exit Sub
    End If

    Set ws = UsersSheet()

    ' Don't let the last active admin be demoted or deactivated
    willBeActiveAdmin = (cmbRole.Value = "Admin" And chkActive.Value)
    If ws.Cells(r, COL_ROLE).Value = "Admin" And ws.Cells(r, COL_ACTIVE).Value = "Yes" _
       And Not willBeActiveAdmin And CountActiveAdmins() <= 1 Then
        MsgBox "You can't demote or deactivate the last active Admin.", vbExclamation, "Update User"
        Exit Sub
    End If

    ' Only change the password if a new one was typed
    If txtPassword.Value <> "" Then
        ws.Cells(r, COL_PASS).Value = HashPassword(uname, txtPassword.Value)
    End If
    ws.Cells(r, COL_ROLE).Value = cmbRole.Value
    ws.Cells(r, COL_ACTIVE).Value = IIf(chkActive.Value, "Yes", "No")

    LoadUsers
    ClearFields
    MsgBox "User updated.", vbInformation, "Update User"
    Exit Sub

ErrHandler:
    MsgBox "Error updating user: " & Err.Number & " - " & Err.Description, vbExclamation, "Update User"
End Sub

' ---------------------------------------------------------
' Delete
' ---------------------------------------------------------
Private Sub cmdDelete_Click()
    Dim ws As Worksheet, r As Long, uname As String

    On Error GoTo ErrHandler

    If lstUsers.SelectedItem Is Nothing Then
        MsgBox "Please select a user.", vbExclamation, "Delete User"
        Exit Sub
    End If

    uname = lstUsers.SelectedItem.Text

    If StrComp(uname, gCurrentUser, vbTextCompare) = 0 Then
        MsgBox "You can't delete the account you are logged in with.", vbExclamation, "Delete User"
        Exit Sub
    End If

    r = FindUserRow(uname)
    If r = 0 Then
        LoadUsers
        Exit Sub
    End If

    Set ws = UsersSheet()

    If ws.Cells(r, COL_ROLE).Value = "Admin" And ws.Cells(r, COL_ACTIVE).Value = "Yes" _
       And CountActiveAdmins() <= 1 Then
        MsgBox "You can't delete the last active Admin.", vbExclamation, "Delete User"
        Exit Sub
    End If

    If MsgBox("Delete user '" & uname & "'? This cannot be undone.", _
              vbYesNo + vbQuestion, "Delete User") <> vbYes Then Exit Sub

    ws.rows(r).Delete

    LoadUsers
    ClearFields
    MsgBox "User deleted.", vbInformation, "Delete User"
    Exit Sub

ErrHandler:
    MsgBox "Error deleting user: " & Err.Number & " - " & Err.Description, vbExclamation, "Delete User"
End Sub
