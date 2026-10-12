Attribute VB_Name = "modStartup"
Option Explicit

' False = Excel stays hidden and only the forms show
Private Const DEV_MODE As Boolean = False

' True if the user has other (visible) workbooks open besides this one
Private Function OtherWorkbooksOpen() As Boolean
    Dim wb As Workbook
    For Each wb In Workbooks
        If Not wb Is ThisWorkbook Then
            If wb.Windows.Count > 0 Then
                If wb.Windows(1).Visible Then
                    OtherWorkbooksOpen = True
                    Exit Function
                End If
            End If
        End If
    Next wb
End Function

Public Sub ShowLoginForm()
    On Error GoTo ErrHandler

    If Not DEV_MODE Then
        If OtherWorkbooksOpen() Then
            ThisWorkbook.Windows(1).Visible = False   ' hide only this workbook
        Else
            Application.Visible = False               ' hide Excel completely
        End If
    End If

    frmLogin.Show vbModal

Cleanup:
    If gExiting Then Exit Sub
    On Error Resume Next
    Application.Visible = True
    ThisWorkbook.Windows(1).Visible = True
    Exit Sub

ErrHandler:
    MsgBox "Error opening the login form: " & Err.Number & " - " & Err.Description, _
           vbCritical, "Startup"
    Resume Cleanup
End Sub

Public Sub ExitApplication()
    Dim i As Long

    gExiting = True
    LogSecurityEvent gUsername, "Exit application"

    On Error Resume Next
    Application.DisplayAlerts = False       ' no save prompts
    ThisWorkbook.Save
    For i = UserForms.Count - 1 To 0 Step -1
        Unload UserForms(i)
    Next i
    On Error GoTo 0

    ' Runs after the form code ends, so Excel isn't destroyed mid-call
    Application.OnTime Now + TimeSerial(0, 0, 1), "ShutDownApp"
End Sub

Public Sub ShutDownApp()
    Dim wb As Workbook, othersDirty As Boolean

    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True

    ' Does any OTHER workbook have unsaved changes?
    For Each wb In Workbooks
        If Not wb Is ThisWorkbook Then
            If Not wb.Saved Then othersDirty = True
        End If
    Next wb

    If OtherWorkbooksOpen() Or othersDirty Then
        ' Don't quit Excel: close only this workbook
        Application.Visible = True
        Application.DisplayAlerts = True
        ThisWorkbook.Close saveChanges:=False   ' already saved in ExitApplication
    Else
        Application.Quit
    End If
End Sub

Public Sub ReturnToApp()
    frmWorkstation.Show
End Sub

