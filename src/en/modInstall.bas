Attribute VB_Name = "modInstall"
Option Explicit

Private Function GetSrcPath() As String
    GetSrcPath = ThisWorkbook.Path & "\PERSONAL.XLSB"
End Function

Private Function GetDstDir() As String
    GetDstDir = Environ$("APPDATA") & "\Microsoft\Excel\XLSTART"
End Function

Private Function GetDstPath() As String
    GetDstPath = GetDstDir() & "\PERSONAL.XLSB"
End Function

'--- backup folder : one level above XLSTART ---
'    It must not go inside XLSTART - Excel opens every file in that
'    folder on startup, backups included.
Private Function GetBackupDir() As String
    Dim s As String
    s = Environ$("APPDATA") & "\Microsoft\Excel\"
    If Len(Dir(s, vbDirectory)) = 0 Then s = Environ$("APPDATA") & "\"
    GetBackupDir = s
End Function

Private Sub EnsureDir(ByVal targetDir As String)
    Dim parentDir As String
    Dim p As Long

    If Len(targetDir) = 0 Then Exit Sub
    If Len(Dir(targetDir, vbDirectory)) > 0 Then Exit Sub

    If Right$(targetDir, 1) = "\" Then targetDir = Left$(targetDir, Len(targetDir) - 1)

    ' stop at the drive root (C:) - nothing above that gets created
    If Len(targetDir) <= 2 Then Exit Sub

    ' for a UNC path (\\server\share) the server and share levels are left alone
    If Left$(targetDir, 2) = "\\" Then
        If UBound(Split(Mid$(targetDir, 3), "\")) < 2 Then Exit Sub
    End If

    p = InStrRev(targetDir, "\")
    If p > 1 Then
        parentDir = Left$(targetDir, p - 1)
        If Len(parentDir) > 0 Then EnsureDir parentDir
    End If

    On Error Resume Next
    MkDir targetDir
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub ClosePersonalIfOpen()
    Dim wb As Workbook
    Dim originalAlerts As Boolean

    On Error Resume Next
    Set wb = Application.Workbooks("PERSONAL.XLSB")
    Err.Clear
    On Error GoTo 0

    If wb Is Nothing Then Exit Sub
    If StrComp(wb.FullName, ThisWorkbook.FullName, vbTextCompare) = 0 Then Exit Sub

    originalAlerts = Application.DisplayAlerts
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = originalAlerts
End Sub

Public Sub PERSONAL_Install()
    Dim srcPath As String
    Dim dstDir As String
    Dim dstPath As String
    Dim originalAlerts As Boolean
    Dim startTime As Double
    Dim bakPath As String
    Dim bakDir As String

    On Error GoTo ErrHandler

    srcPath = GetSrcPath()
    dstDir = GetDstDir()
    dstPath = GetDstPath()

    originalAlerts = Application.DisplayAlerts

    ' 1. check the source
    If Len(Dir(srcPath)) = 0 Then
        MsgBox "PERSONAL.XLSB is not in the same folder as this file." & vbCrLf & vbCrLf & _
               "This folder: " & ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "Put both files in one folder and run it again.", vbExclamation, "Install failed"
        Exit Sub
    End If

    If FileLen(srcPath) = 0 Then
        MsgBox "The PERSONAL.XLSB source file looks damaged (size 0)." & vbCrLf & _
               "Stopping here.", vbExclamation, "Install failed"
        Exit Sub
    End If

    ' 2. make sure the target folder is there
    EnsureDir dstDir
    If Len(Dir(dstDir, vbDirectory)) = 0 Then
        MsgBox "The target folder could not be created." & vbCrLf & dstDir, vbExclamation, "Install failed"
        Exit Sub
    End If

    ' 3. close PERSONAL.XLSB if it is open
    ClosePersonalIfOpen

    ' 4. back up whatever is already there
    '    It must not go inside XLSTART - Excel opens every file in that folder.
    '    So it goes one level up (%APPDATA%\Microsoft\Excel).
    If Len(Dir(dstPath)) > 0 Then
        bakDir = GetBackupDir()
        bakPath = bakDir & "PERSONAL_backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".XLSB"
        On Error Resume Next
        FileCopy dstPath, bakPath
        Err.Clear
        On Error GoTo ErrHandler
        If Len(Dir(bakPath)) = 0 Then bakPath = ""
    End If

    ' 5. remove the old file
    If Len(Dir(dstPath)) > 0 Then
        On Error Resume Next
        SetAttr dstPath, vbNormal
        Kill dstPath
        Err.Clear
        On Error GoTo ErrHandler

        startTime = Timer
        Do While Len(Dir(dstPath)) > 0
            DoEvents
            If Timer < startTime Then Exit Do
            If Timer - startTime >= 1 Then Exit Do
        Loop

        If Len(Dir(dstPath)) > 0 Then
            MsgBox "PERSONAL.XLSB may still be in use by another Excel window." & vbCrLf & vbCrLf & _
                   "Close every Excel window and install again.", vbExclamation, "Install failed"
            Application.DisplayAlerts = originalAlerts
            Exit Sub
        End If
    End If

    ' 6. copy
    FileCopy srcPath, dstPath

    ' 7. verify
    If Len(Dir(dstPath)) = 0 Then
        MsgBox "Copying the file failed.", vbExclamation, "Install failed"
        Application.DisplayAlerts = originalAlerts
        Exit Sub
    End If

    If FileLen(srcPath) <> FileLen(dstPath) Or FileLen(dstPath) = 0 Then
        On Error Resume Next
        Kill dstPath
        Err.Clear
        On Error GoTo ErrHandler
        MsgBox "The copied file did not verify. Stopping here.", vbExclamation, "Install failed"
        Application.DisplayAlerts = originalAlerts
        Exit Sub
    End If

    ' 8. done
    Application.DisplayAlerts = originalAlerts
    MsgBox "Installation finished." & vbCrLf & vbCrLf & _
           "Close every Excel window, then start Excel again." & vbCrLf & _
           "The work tools load from the next start." & vbCrLf & vbCrLf & _
           "Installed at: " & dstPath & _
           IIf(Len(bakPath) > 0, vbCrLf & vbCrLf & _
               "You already had a PERSONAL.XLSB, so it was backed up." & vbCrLf & _
               bakPath, ""), vbInformation, "Install complete"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = originalAlerts
    MsgBox "An error occurred while installing." & vbCrLf & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Error text: " & Err.Description, vbCritical, "Install error"
End Sub

Public Sub PERSONAL_Remove()
    Dim dstPath As String
    Dim ans As VbMsgBoxResult
    Dim startTime As Single

    On Error GoTo ErrHandler

    dstPath = GetDstPath()

    If Len(Dir(dstPath)) = 0 Then
        MsgBox "There is no PERSONAL.XLSB installed.", vbInformation, "Remove"
        Exit Sub
    End If

    ans = MsgBox("Delete the installed PERSONAL.XLSB?" & vbCrLf & vbCrLf & dstPath, _
                 vbYesNo + vbQuestion, "Confirm removal")
    If ans <> vbYes Then Exit Sub

    ' close it first if this Excel has it open
    ClosePersonalIfOpen

    On Error Resume Next
    SetAttr dstPath, vbNormal
    Kill dstPath
    Err.Clear
    On Error GoTo ErrHandler

    ' the file handle can take a moment to let go, so wait briefly
    startTime = Timer
    Do While Len(Dir(dstPath)) > 0
        DoEvents
        If Timer < startTime Then Exit Do
        If Timer - startTime >= 1 Then Exit Do
        On Error Resume Next
        Kill dstPath
        Err.Clear
        On Error GoTo ErrHandler
    Loop

    If Len(Dir(dstPath)) > 0 Then
        MsgBox "PERSONAL.XLSB may still be in use by another Excel window." & vbCrLf & vbCrLf & _
               "Close every Excel window and delete it again.", vbExclamation, "Remove failed"
        Exit Sub
    End If

    MsgBox "Removal finished." & vbCrLf & vbCrLf & _
           "Close every Excel window and start Excel again," & vbCrLf & _
           "and the work tools shortcuts will be gone.", vbInformation, "Remove complete"
    Exit Sub

ErrHandler:
    MsgBox "An error occurred while removing." & vbCrLf & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Error text: " & Err.Description, vbCritical, "Remove error"
End Sub
