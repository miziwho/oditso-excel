Attribute VB_Name = "mod_WorkTools"
Option Explicit

'============================================================
'  Excel Work Tools  v1.1.0   (2026-11-01)
'  A single module, for PERSONAL.XLSB only
'
'  Install : in the VBE, use [Import File] to add this one file,
'         then paste ThisWorkbook_code.txt into ThisWorkbook. Done.
'  Settings : kept in PERSONAL.XLSB defined names, with no sheet.
'
'  Shortcuts
'    Ctrl+Shift+Q  Copy Sheet to New Workbook
'    Ctrl+Shift+X  Comma Format
'    Ctrl+Shift+W  AutoFit Rows / Columns
'    Ctrl+Shift+S  Freeze / Unfreeze Header Row
'    Ctrl+Shift+Y  Text to Columns - as General
'    Ctrl+Shift+I  Text to Columns - as Text
'    Ctrl+Shift+H  Tidy All Sheets
'    Ctrl+Shift+J  Create PivotTable (new sheet)
'    Ctrl+Shift+T  Classic PivotTable Layout
'    Ctrl+Shift+D  Highlight Duplicate Rows
'    Ctrl+Shift+V  Duplicates by Key Column (highlight · group no. · extract)
'    Ctrl+Shift+C  Clear Cell Fill
'    Ctrl+Shift+U  Table Borders
'    Ctrl+Shift+A  Filter (toggle AutoFilter / clear Advanced Filter)
'    Ctrl+Shift+F  Run Advanced Filter
'    Ctrl+Shift+R  Set Up Advanced Filter Criteria Columns
'    Ctrl+Shift+G  SUBTOTAL Sum
'    Ctrl+Alt+G    SUBTOTAL Count (filled cells)
'    Ctrl+Shift+E  Export Sheet as Text
'    Ctrl+Shift+N  Copy File Path
'    Ctrl+Shift+Z  Filter by Selected Cells
'    Ctrl+Shift+P  Set Password
'    Ctrl+Shift+O  Open New Window
'    Ctrl+Shift+K  Paste / Remove Shortcut List
'    Ctrl+Shift+L  Multi VLOOKUP (spread several hits sideways)
'    Ctrl+Alt+F    Find by Value (matches the result)
'    Ctrl+Shift+1  Fill Formula Down
'    Ctrl+Shift+2  Insert Index Column
'    Ctrl+Shift+3  Select to End of Table (ignores blanks)
'    Ctrl+Shift+4  Fill Blanks From Above
'    Ctrl+Shift+M  Work Tools Menu
'
'    Ctrl+Alt+Z    Center / Default (toggle)
'    Ctrl+Alt+S    Center Across Selection (without merging)
'    Ctrl+Alt+A    Flash Fill (an upgrade on Ctrl+E, to the end of the table)
'    Ctrl+Alt+Q    Auto Fill Series (1,2,3 / Mon,Tue,Wed, to the end of the table)
'    Ctrl+Alt+E    Remove Rows Above Header
'    Ctrl+Alt+X    Copy / Paste Formatting (one key, alternating)
'    Ctrl+Alt+D    Compare Two Columns (both ways, is it in the list)
'    Ctrl+Alt+B    Column Headers to Row Data (unpivot)
'    (automatic)   asks about the open password when a saved file is closed
'============================================================

Private Const FORMAT_BRUSH_SECONDS As Long = 40
Public Const TOOL_VERSION As String = "1.1.0"
Public Const TOOL_VERSION_DATE As String = "2026-11-01"
Public Const BUILD_TIME As String = "2026-11-01 18:50"
Public Const AUTHOR As String = "SH"
Public Const SITE As String = "oditso.com"

Private Const DEFAULT_PASSWORD As String = "1234"
Private Const PASSWORD_MAX_LEN As Long = 15
Private Const BAR_MAIN_MENU As String = "WorkToolsBar"
Private Const CHEAT_SHEET_SHAPE As String = "WorkTools_ShortcutList"
Private Const BAR_COLUMN_PICK As String = "ColumnPickMenu"
Private Const BAR_MULTI_LOOKUP As String = "MultiLookupMenu"
Private Const BAR_COMPARE_COLS As String = "CompareColumnsMenu"
Private Const BAR_TWO_SPOT As String = "TwoSpotMenu"
Private Const BAR_COMMA_FORMAT As String = "CommaFormatBar"
Private Const BAR_UNPIVOT As String = "UnpivotBar"

Private mBusy As Boolean               ' keeps it from re-entering
Private mSuppressOnce As Boolean       ' keeps it from asking again during the delayed close
Private mPendingCloseName As String    ' the workbook to close later

' kept so the SUBTOTAL sum can be undone
Private mUndoWb As String
Private mUndoWs As String
Private mUndoAddr As String
Private mUndoFormula As Variant
Private mUndoBold As Variant
Private mUndoNF As Variant
Private mUndoFunc As Long
Private mUndoTitle As String

' kept so the duplicate highlighting can be undone
Private mDupWb As String
Private mDupWs As String
Private mDupAddr As String
Private mDupRows As Variant
Private mDupColor As Variant
Private mDupIndex As Variant
Private mDupCount As Long

' kept so the shortcut list can be undone
Private mCheatWb As String
Private mCheatWs As String

' kept so Text to Columns can be undone
Private mTxtWb As String
Private mTxtWs As String
Private mTxtAddr As String
Private mTxtValue As Variant
Private mTxtFormat As Variant
Private mTxtOn As Boolean

' kept so comma format / AutoFit / criteria columns can be undone
Private mFmtWb As String, mFmtWs As String, mFmtAddr As String
Private mFmtValue As Variant, mFmtOn As Boolean

Private mFitWb As String, mFitWs As String, mFitAddr As String
Private mFitColW As Variant, mFitRowH As Variant, mFitOn As Boolean

Private mCondWb As String, mCondWs As String, mCondAddr As String
Private mCondF As Variant, mCondN As Variant
Private mCondC As Variant, mCondI As Variant
Private mCondB As Variant, mCondFC As Variant
Private mCondOn As Boolean

' kept so the table borders can be undone
Private mBdrWb As String, mBdrWs As String, mBdrAddr As String
Private mBdrLine As Variant, mBdrBold As Variant, mBdrAlign As Variant
Private mBdrOn As Boolean

' kept so the cleared cell fill can be undone
Private mBgWb As String, mBgWs As String, mBgAddr As String
Private mBgColor As Variant, mBgIndex As Variant, mBgOn As Boolean

' stops each step from reporting when the combined button calls them
Private mQuiet As Boolean
Private mDupGrp As Long, mDupHit As Long

' duplicates by key column - which columns are picked
Private mColPick() As Boolean
Private mColWs As String
Private mColAddr As String

' multi VLOOKUP - which columns are picked
Private mLkKey() As Boolean
Private mLkVal() As Boolean
Private mLkWs As String
Private mLkAddr As String
Private mLkFormula As Boolean
Private mLkDedup As Boolean

' column headers to row data - which columns are picked
'   mUpVal(i) = True  ->  that column moves down into rows (+)
'   False            ->  it stays where it is (o)
Private mUpVal() As Boolean
Private mUpWs As String
Private mUpAddr As String
Private mUpBlank As Boolean
Private mUpDown As Boolean
Private mUpName1 As String
Private mUpName2 As String

' compare two columns - which columns are picked
Private mCmpA As Long
Private mCmpB As Long
Private mCmpWs As String
Private mCmpAddr As String
Private mCmpFormula As Boolean

' kept so removing the rows above the header can be undone
Private mDelWb As String, mDelWs As String
Private mDelRows As Long, mDelCols As Long
Private mDelF As Variant, mDelN As Variant
Private mDelC As Variant, mDelI As Variant
Private mDelB As Variant, mDelA As Variant
Private mDelH As Variant
Private mDelOn As Boolean

' the copied formatting
Private mFmtSrcWb As String, mFmtSrcWs As String, mFmtSrcAddr As String
Private mFmtSrcOn As Boolean
Private mFmtSrcTime As Date

' kept so comparing the same rows can be undone
Private mDifWb As String, mDifWs As String
Private mDifResAddr As String, mDifColAddr As String
Private mDifRes As Variant
Private mDifColor As Variant, mDifIndex As Variant
Private mDifOn As Boolean
Private mDifIns As Boolean
Private mDifInsCol As Long

' the two spots picked with Ctrl
Private mCmpDirWs As String
Private mCmpDirA As String, mCmpDirB As String

' kept so the alignment can be undone
Private mAlignWb As String, mAlignWs As String, mAlignAddr As String
Private mAlignVal As Variant, mAlignOn As Boolean

' kept so the insert can be undone
Private mInsWb As String, mInsWs As String
Private mInsCol As Long, mInsCount As Long, mInsOn As Boolean

' where the SUBTOTAL toggle sits
Private mUndoOn As Boolean


'============================================================
'  A. Saving / reading settings  (defined names - no sheet needed)
'============================================================

Private Function ReadSetting(ByVal key2 As String, ByVal defaultVal As String) As String
    Dim s As String

    On Error Resume Next
    s = CStr(ThisWorkbook.Names("cfg_" & key2).RefersTo)
    Err.Clear
    On Error GoTo 0

    ' RefersTo is stored in the form  ="value"
    If Len(s) >= 3 Then
        If Left$(s, 2) = "=""" And Right$(s, 1) = """" Then
            s = Mid$(s, 3, Len(s) - 3)
            s = Replace(s, """""", """")
        Else
            s = ""
        End If
    Else
        s = ""
    End If

    If Len(s) = 0 Then s = defaultVal
    ReadSetting = s
End Function

Private Sub WriteSetting(ByVal key2 As String, ByVal cellVal As String)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="cfg_" & key2, _
                           RefersTo:="=""" & Replace(cellVal, """", """""") & """", _
                           Visible:=False
    ThisWorkbook.Save
    Err.Clear
    On Error GoTo 0
End Sub

Public Function ReadDefaultPassword() As String
    ReadDefaultPassword = ReadSetting("pwd", DEFAULT_PASSWORD)
End Function

Public Function AskOnCloseEnabled() As Boolean
    AskOnCloseEnabled = (ReadSetting("auto", "1") <> "0")
End Function


'============================================================
'  B. Startup / shortcuts
'============================================================

Private Function MacroPrefix() As String
    MacroPrefix = "'" & ThisWorkbook.Name & "'!"
End Function

Public Sub RegisterShortcuts()
    Dim p As String
    p = MacroPrefix()

    On Error Resume Next
    Application.OnKey "^+q", p & "CopyToNewWorkbook"
    Application.OnKey "^+x", p & "CommaFormat"
    Application.OnKey "^+w", p & "AutoFitRowsCols"
    Application.OnKey "^+s", p & "FreezeHeaderRow"
    Application.OnKey "^+y", p & "TextToGeneral"
    Application.OnKey "^+i", p & "TextToText"
    Application.OnKey "^+h", p & "TidyAllSheets"
    Application.OnKey "^+j", p & "CreatePivot"
    Application.OnKey "^+t", p & "ClassicPivotCleanup"
    Application.OnKey "^+d", p & "HighlightDuplicateRows"
    Application.OnKey "^+v", p & "ColumnDuplicates"
    Application.OnKey "^+c", p & "ClearFillColor"
    Application.OnKey "^+u", p & "TableBorders"
    Application.OnKey "^+a", p & "ToggleAutoFilter"
    Application.OnKey "^+f", p & "AdvancedFilter"
    Application.OnKey "^+r", p & "BuildCriteriaBlock"
    Application.OnKey "^+g", p & "SubtotalSum"
    Application.OnKey "^+e", p & "ExportSheetAsText"
    Application.OnKey "^+n", p & "CopyFilePath"
    Application.OnKey "^+z", p & "FilterBySelectedValue"
    Application.OnKey "^+p", p & "SetPassword"
    Application.OnKey "^+o", p & "OpenNewWindow"
    Application.OnKey "^+k", p & "CheatSheet"
    Application.OnKey "^+l", p & "MultiLookup"
    Application.OnKey "^%f", p & "FindByValue"
    Application.OnKey "^+1", p & "FillFormulaDown"
    Application.OnKey "^+2", p & "InsertIndexColumn"
    Application.OnKey "^+3", p & "SelectToEnd"
    Application.OnKey "^+4", p & "FillBlanksFromAbove"
    Application.OnKey "^+m", p & "MacroLauncher"

    Application.OnKey "^%z", p & "ToggleCenterAlign"
    Application.OnKey "^%s", p & "CenterAcrossSelection"
    Application.OnKey "^%a", p & "FlashFillToEnd"
    Application.OnKey "^%q", p & "AutoFillSeries"
    Application.OnKey "^%e", p & "RemoveRowsAboveHeader"
    Application.OnKey "^%x", p & "FormatBrush"
    Application.OnKey "^%d", p & "CompareTwoColumns"
    Application.OnKey "^%b", p & "Unpivot"
    Application.OnKey "^%g", p & "SubtotalCount"
    Err.Clear
    On Error GoTo 0

    mBusy = False
    mSuppressOnce = False
    mPendingCloseName = ""
End Sub

Public Sub UnregisterShortcuts()
    On Error Resume Next
    Application.OnKey "^+q"
    Application.OnKey "^+z"
    Application.OnKey "^+w"
    Application.OnKey "^+s"
    Application.OnKey "^+y"
    Application.OnKey "^+i"
    Application.OnKey "^+h"
    Application.OnKey "^+j"
    Application.OnKey "^+t"
    Application.OnKey "^+d"
    Application.OnKey "^+v"
    Application.OnKey "^+c"
    Application.OnKey "^+u"
    Application.OnKey "^+a"
    Application.OnKey "^+f"
    Application.OnKey "^+g"
    Application.OnKey "^+r"
    Application.OnKey "^+e"
    Application.OnKey "^+n"
    Application.OnKey "^+x"
    Application.OnKey "^+p"
    Application.OnKey "^+o"
    Application.OnKey "^+k"
    Application.OnKey "^+l"
    Application.OnKey "^%f"
    Application.OnKey "^+1"
    Application.OnKey "^+2"
    Application.OnKey "^+3"
    Application.OnKey "^+4"
    Application.OnKey "^+m"
    Application.OnKey "^%z"
    Application.OnKey "^%s"
    Application.OnKey "^%a"
    Application.OnKey "^%q"
    Application.OnKey "^%e"
    Application.OnKey "^%x"
    ClearFormatBrush
    Application.OnKey "^%d"
    Application.OnKey "^%b"
    Application.OnKey "^%g"
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------
' reinitialize by hand when the tools stop working (run from Alt+F8)
'------------------------------------------------------------
Public Sub WorkTools_Restart()
    On Error Resume Next
    ThisWorkbook.ConnectEvents
    Err.Clear
    On Error GoTo 0
    RegisterShortcuts
    MsgBox "Work Tools has been reinitialized." & vbCrLf & _
           "Version " & TOOL_VERSION, vbInformation, "Work Tools"
End Sub

'------------------------------------------------------------
' diagnostics (run from Alt+F8) - to check the state when something goes wrong
'------------------------------------------------------------
Public Sub WorkTools_Diagnostics()
    Dim s As String
    Dim eventState As String

    eventState = "unknown"
    On Error Resume Next
    eventState = IIf(ThisWorkbook.EventsConnected(), "connected", "disconnected")
    Err.Clear
    On Error GoTo 0

    s = "Excel Work Tools " & TOOL_VERSION & "  (build " & BUILD_TIME & ")" & vbCrLf & _
        "Made by " & AUTHOR & " in Korea  ·  " & SITE & vbCrLf & vbCrLf & _
        "Excel version   : " & Application.Version & _
                              " (" & IIf(OfficeBitness() = 64, "64bit", "32bit") & ")" & vbCrLf & _
        "Installed at    : " & ThisWorkbook.FullName & vbCrLf & _
        "Close event     : " & eventState & vbCrLf & _
        "Ask on close    : " & IIf(AskOnCloseEnabled(), "on", "off") & vbCrLf & _
        "Default password : " & ReadDefaultPassword() & vbCrLf & vbCrLf & _
        "If the shortcuts do not work, run [WorkTools_Restart]."

    MsgBox s, vbInformation, "Work Tools Diagnostics"
End Sub

Private Function OfficeBitness() As Long
    #If Win64 Then
        OfficeBitness = 64
    #Else
        OfficeBitness = 32
    #End If
End Function


'============================================================
'  C. Open password engine  (shared by manual and automatic)
'============================================================

'------------------------------------------------------------
' ask for the password   returns 0 = ok / 1 = cancelled
'------------------------------------------------------------
Private Function AskPassword(ByRef Pwd As String) As Long
    Dim s1 As String, s2 As String
    Dim def As String
    Dim r As VbMsgBoxResult

    Pwd = ""
    def = ReadDefaultPassword()

    Do
        s1 = InputBox("Type the open password to set. (1-" & PASSWORD_MAX_LEN & " chars)", _
                      "Set Open Password", def)

        If StrPtr(s1) = 0 Then                          ' the [Cancel] button
            AskPassword = 1
            Exit Function
        End If

        If Len(s1) = 0 Then
            r = MsgBox("An empty password cannot be used." & vbCrLf & _
                       "Do you want to enter it again?", vbOKCancel + vbExclamation, "Set Open Password")
            If r = vbCancel Then
                AskPassword = 1
                Exit Function
            End If
        ElseIf Len(s1) > PASSWORD_MAX_LEN Then
            MsgBox "The password must be " & PASSWORD_MAX_LEN & " characters or fewer.", _
                   vbExclamation, "Set Open Password"
        Else
            Exit Do
        End If
    Loop

    ' skip the confirmation when the default password is used as is
    If StrComp(s1, def, vbBinaryCompare) = 0 Then
        Pwd = s1
        AskPassword = 0
        Exit Function
    End If

    Do
        s2 = InputBox("Please type the same password once more.", "Password Check")

        If StrPtr(s2) = 0 Then
            AskPassword = 1
            Exit Function
        End If

        If StrComp(s1, s2, vbBinaryCompare) = 0 Then Exit Do

        r = MsgBox("The two passwords do not match. (Case sensitive.)" & vbCrLf & _
                   "Do you want to enter it again?", vbOKCancel + vbExclamation, "Password Check")
        If r = vbCancel Then
            AskPassword = 1
            Exit Function
        End If
    Loop

    Pwd = s1
    AskPassword = 0
End Function

'------------------------------------------------------------
' write the open password into the file itself
'   SaveAs to the same path in the same format, then verify the result.
'------------------------------------------------------------
Private Function ApplyOpenPassword(ByVal Wb As Workbook, _
                              ByVal Pwd As String, _
                              ByRef ErrMsg As String) As Boolean

    Dim oldAlerts As Boolean, oldEvents As Boolean, oldScreen As Boolean
    Dim sFull As String
    Dim lFmt As Long
    Dim ok As Boolean
    Dim fileOK As Boolean

    ErrMsg = ""

    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating

    On Error GoTo Fail

    If Wb Is Nothing Then
        ErrMsg = "Could not find the target workbook."
        GoTo Restore
    End If
    If Wb.Path = "" Then
        ErrMsg = "This file has never been saved."
        GoTo Restore
    End If

    sFull = Wb.FullName
    lFmt = Wb.FileFormat

    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Wb.SaveAs Filename:=sFull, FileFormat:=lFmt, Password:=Pwd, CreateBackup:=False

    '--- verify the save ---
    ok = Wb.HasPassword

    fileOK = True
    On Error Resume Next
    fileOK = (Len(Dir$(sFull)) > 0)
    If Err.Number <> 0 Then
        fileOK = True                       ' if Dir itself fails, leave the verdict open
        Err.Clear
    End If
    On Error GoTo Fail

    If Not fileOK Then ok = False
    If Not ok Then ErrMsg = "The file was saved but the open password could not be confirmed."

Restore:
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    ApplyOpenPassword = ok
    Exit Function

Fail:
    ErrMsg = "Error number: " & Err.Number & vbCrLf & "Description: " & Err.Description
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    ApplyOpenPassword = False
End Function

'------------------------------------------------------------
' set password (Ctrl+Shift+P) - run by hand
'------------------------------------------------------------
Public Sub SetPassword()
    Dim Wb As Workbook
    Dim Pwd As String
    Dim ErrMsg As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "Open the file you want to protect first, then run this.", vbExclamation, "Set Open Password"
        Exit Sub
    End If

    Set Wb = ActiveWorkbook

    If StrComp(Wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then
        MsgBox "Open the file you want to protect first, then run this.", vbExclamation, "Set Open Password"
        Exit Sub
    End If
    If Wb.Path = "" Then
        MsgBox "This file has not been saved yet." & vbCrLf & _
               "Please save the file first, then run again.", vbExclamation, "Set Open Password"
        Exit Sub
    End If
    If Wb.ReadOnly Then
        MsgBox "A password cannot be set on a read-only file.", _
               vbExclamation, "Set Open Password"
        Exit Sub
    End If
    If Wb.HasPassword Then
        MsgBox "This file already has an open password.", _
               vbInformation, "Set Open Password"
        Exit Sub
    End If

    If AskPassword(Pwd) <> 0 Then Exit Sub

    If ApplyOpenPassword(Wb, Pwd, ErrMsg) Then
        MsgBox "Open password set and the file was saved." & vbCrLf & vbCrLf & Wb.Name, _
               vbInformation, "Set Open Password"
    Else
        MsgBox "An error occurred while setting the password or saving." & vbCrLf & vbCrLf & ErrMsg, _
               vbCritical, "Set Open Password"
    End If
End Sub


'============================================================
'  D. Asking about the password on close
'============================================================

Public Sub ToggleAskOnClose()
    Dim current As Boolean
    current = AskOnCloseEnabled()
    WriteSetting "auto", IIf(current, "0", "1")

    If current Then
        MsgBox "Asking for a password on close is now off." & vbCrLf & _
               "You can turn it back on from the same menu.", vbInformation, "Work Tools"
    Else
        MsgBox "You will be asked again about setting a password when closing a saved file.", _
               vbInformation, "Work Tools"
    End If
End Sub

Public Sub ChangeDefaultPassword()
    Dim s As String

    s = InputBox("This password is pre-filled in the password box. (1-" & PASSWORD_MAX_LEN & " chars)", _
                 "Change Default Password", ReadDefaultPassword())

    If StrPtr(s) = 0 Then Exit Sub
    If Len(s) = 0 Then
        MsgBox "An empty value cannot be used.", vbExclamation, "Change Default Password"
        Exit Sub
    End If
    If Len(s) > PASSWORD_MAX_LEN Then
        MsgBox "The password must be " & PASSWORD_MAX_LEN & " characters or fewer.", vbExclamation, "Change Default Password"
        Exit Sub
    End If

    WriteSetting "pwd", s
    MsgBox "The default password is now '" & s & "'.", vbInformation, "Change Default Password"
End Sub

'------------------------------------------------------------
' decide whether this workbook should be checked
'------------------------------------------------------------
Private Function IsTargetWorkbook(ByVal Wb As Workbook) As Boolean
    Dim ext As String
    Dim p As Long

    On Error GoTo Fail

    If Wb Is Nothing Then Exit Function
    If StrComp(Wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then Exit Function
    If Wb.Path = "" Then Exit Function                  ' a new workbook that has never been saved
    If Wb.IsAddin Then Exit Function
    If Wb.ReadOnly Then Exit Function

    p = InStrRev(Wb.Name, ".")
    If p = 0 Then Exit Function
    ext = LCase$(Mid$(Wb.Name, p + 1))

    Select Case ext
        Case "xlsx", "xlsm", "xlsb", "xls"
            IsTargetWorkbook = True
    End Select

    Exit Function
Fail:
    IsTargetWorkbook = False
End Function

'------------------------------------------------------------
' runs just before close  (called from App_WorkbookBeforeClose in ThisWorkbook)
'------------------------------------------------------------
Public Sub AskPasswordBeforeClose(ByVal Wb As Workbook, ByRef Cancel As Boolean)
    Dim ans As VbMsgBoxResult
    Dim Pwd As String
    Dim ErrMsg As String
    Dim msg As String

    On Error GoTo SafeExit

    If mBusy Then Exit Sub
    If mSuppressOnce Then Exit Sub
    If Not AskOnCloseEnabled() Then Exit Sub
    If Not IsTargetWorkbook(Wb) Then Exit Sub
    If Wb.HasPassword Then Exit Sub

    msg = "This file has no open password set." & vbCrLf & vbCrLf & _
          "  " & Wb.Name & vbCrLf & vbCrLf & _
          "[Yes]     Set a password and close." & vbCrLf & _
          "[No]      Close without a password." & vbCrLf & _
          "[Cancel]  Leave it open.  (Esc)" & vbCrLf & vbCrLf & _
          "Change the default : Ctrl+Shift+M -> turn off"

    ans = MsgBox(msg, vbYesNoCancel + vbQuestion, "Open Password Check")

    If ans = vbNo Then Exit Sub                         ' leave the close as it was

    If ans = vbCancel Then                              ' cancel the close itself
        Cancel = True
        Exit Sub
    End If

    '--- [Yes] : hold the close for a moment and set the password ---
    Cancel = True
    mBusy = True

    If AskPassword(Pwd) <> 0 Then
        mBusy = False
        Exit Sub                                        ' keep the file open
    End If

    If Not ApplyOpenPassword(Wb, Pwd, ErrMsg) Then
        mBusy = False
        MsgBox "An error occurred while setting the password or saving." & vbCrLf & vbCrLf & _
               ErrMsg & vbCrLf & vbCrLf & _
               "The file was not closed. Please check it and try again.", _
               vbCritical, "Open Password Check"
        Exit Sub                                        ' never close it on failure
    End If

    MsgBox "Open password set and the file was saved.", vbInformation, "Open Password Check"

    mPendingCloseName = Wb.Name
    mBusy = False

    ' close it safely once the event is over
    On Error Resume Next
    Application.OnTime Now, MacroPrefix() & "DelayedClose"
    Err.Clear
    On Error GoTo 0
    Exit Sub

SafeExit:
    mBusy = False
    Cancel = True
    MsgBox "An error occurred while checking the password." & vbCrLf & vbCrLf & _
           "Error number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & vbCrLf & _
           "The file was not closed.", vbCritical, "Open Password Check"
End Sub

'------------------------------------------------------------
' delayed close (Application.OnTime callback)
'------------------------------------------------------------
Public Sub DelayedClose()
    Dim Wb As Workbook
    Dim nm As String

    nm = mPendingCloseName
    mPendingCloseName = ""
    If Len(nm) = 0 Then Exit Sub

    On Error Resume Next
    Set Wb = Application.Workbooks(nm)
    Err.Clear
    On Error GoTo 0
    If Wb Is Nothing Then Exit Sub

    mSuppressOnce = True
    On Error Resume Next
    Wb.Close SaveChanges:=False
    Err.Clear
    On Error GoTo 0
    mSuppressOnce = False
End Sub

'------------------------------------------------------------
' clear the status bar (Application.OnTime callback)
'------------------------------------------------------------
Public Sub ClearStatusBar()
    On Error Resume Next
    Application.StatusBar = False
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub ScheduleStatusClear()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 4), MacroPrefix() & "ClearStatusBar"
    Err.Clear
    On Error GoTo 0
End Sub


'============================================================
'  E. Work tools
'============================================================

'------------------------------------------------------------
' 1. CopyToNewWorkbook (Ctrl+Shift+Q)
'------------------------------------------------------------
Public Sub CopyToNewWorkbook()
    Dim sh As Object

    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then
        MsgBox "There is no sheet to copy.", vbExclamation, "Copy to New Workbook"
        Exit Sub
    End If

    Set sh = ActiveSheet
    If sh Is Nothing Then Exit Sub

    sh.Copy
    Exit Sub

Fail:
    MsgBox "Could not copy to a new workbook." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Copy to New Workbook"
End Sub

'------------------------------------------------------------
' 2. CommaFormat (Ctrl+Shift+X)
'------------------------------------------------------------
Public Sub CommaFormat()
    On Error GoTo Fail

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell first.", vbExclamation, "Comma Format"
        Exit Sub
    End If

    ' kept so it can be undone
    On Error Resume Next
    If Selection.Cells.CountLarge <= 100000 Then
        mFmtWb = Selection.Worksheet.Parent.Name
        mFmtWs = Selection.Worksheet.Name
        mFmtAddr = Selection.Address
        mFmtValue = Selection.NumberFormat
        mFmtOn = True
    Else
        mFmtOn = False
    End If
    Err.Clear
    On Error GoTo Fail

    Selection.NumberFormat = CommaFormatString()

    If mFmtOn Then
        On Error Resume Next
        Application.OnUndo "Undo Comma Format", MacroPrefix() & "UndoCommaFormat"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "Comma format applied." & _
                            IIf(mFmtOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    MsgBox "Could not apply the format." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comma Format"
End Sub

'------------------------------------------------------------
' 3. ClassicPivot (Ctrl+Shift+T)
'------------------------------------------------------------
Public Sub ClassicPivotCleanup()
    Dim errNo As Long, errDesc As String
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim i As Long
    Dim oldScreen As Boolean

    oldScreen = Application.ScreenUpdating

    On Error Resume Next
    Set pt = ActiveCell.PivotTable
    Err.Clear
    On Error GoTo 0

    If pt Is Nothing Then
        MsgBox "Select a cell inside the PivotTable, then run again.", vbExclamation, "Classic PivotTable Cleanup"
        Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False

    With pt
        .InGridDropZones = True                 ' classic PivotTable layout
        .RowAxisLayout xlTabularRow             ' each row field in its own column
    End With

    pt.ManualUpdate = True                      ' so it does not recalculate on every change

    For Each pf In pt.RowFields                 ' remove the subtotals on every row field
        For i = 1 To 12
            On Error Resume Next
            pf.Subtotals(i) = False
            Err.Clear
            On Error GoTo Fail
        Next i
    Next pf

    pt.RepeatAllLabels xlRepeatLabels           ' repeat every item label

    On Error Resume Next
    pt.ManualUpdate = False                     ' back to refreshing on its own
    Err.Clear
    On Error GoTo Fail

    Application.ScreenUpdating = oldScreen

    MsgBox "PivotTable ready!" & vbCrLf & _
           "· Show in classic layout" & vbCrLf & _
           "· Put row fields in separate columns" & vbCrLf & _
           "· Remove all subtotals" & vbCrLf & _
           "· Repeat all item labels" & vbCrLf & vbCrLf & _
           "* This cannot be undone with Ctrl+Z.", vbInformation, "Classic PivotTable Cleanup"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not pt Is Nothing Then pt.ManualUpdate = False
    Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = oldScreen
    MsgBox "An error occurred while applying the setting." & vbCrLf & _
           "Error number: " & errNo & vbCrLf & errDesc, _
           vbExclamation, "Classic PivotTable Cleanup"
End Sub

'------------------------------------------------------------
' 4. TidyAllSheets (Ctrl+Shift+H)
'------------------------------------------------------------
Public Sub TidyAllSheets()
    Dim Wb As Workbook
    Dim ws As Worksheet
    Dim wsFirst As Worksheet
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim n As Long

    If ActiveWorkbook Is Nothing Then Exit Sub
    Set Wb = ActiveWorkbook

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each ws In Wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If wsFirst Is Nothing Then Set wsFirst = ws
            On Error Resume Next
            ws.Activate
            ActiveWindow.Zoom = 100
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
            ws.Range("A1").Select
            Err.Clear
            On Error GoTo Fail
            n = n + 1
        End If
    Next ws

    If Not wsFirst Is Nothing Then
        On Error Resume Next
        wsFirst.Activate
        Err.Clear
        On Error GoTo Fail
    End If

    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents

    Application.StatusBar = "Sheet " & n & " cleaned up."
    ScheduleStatusClear
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    MsgBox "Could not tidy the sheets." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Tidy All Sheets"
End Sub

'------------------------------------------------------------
' 5. AutoFitRowsCols (Ctrl+Shift+W)
'------------------------------------------------------------
Public Sub AutoFitRowsCols()
    Dim rng As Range
    Dim oldScreen As Boolean
    Dim i As Long
    Dim prevWidth() As Variant
    Dim prevHeight() As Variant

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell first.", vbExclamation, "AutoFit Rows and Columns"
        Exit Sub
    End If

    If Selection.Cells.CountLarge >= 2 Then
        Set rng = Selection
    Else
        Set rng = ActiveCell.CurrentRegion       ' fit the total row width as well
    End If

    oldScreen = Application.ScreenUpdating

    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' kept so it can be undone
    On Error Resume Next
    ReDim prevWidth(1 To rng.Columns.Count)
    ReDim prevHeight(1 To rng.Rows.Count)
    For i = 1 To rng.Columns.Count
        prevWidth(i) = rng.Columns(i).ColumnWidth
    Next i
    For i = 1 To rng.Rows.Count
        prevHeight(i) = rng.Rows(i).RowHeight
    Next i
    mFitWb = rng.Worksheet.Parent.Name
    mFitWs = rng.Worksheet.Name
    mFitAddr = rng.Address
    mFitColW = prevWidth
    mFitRowH = prevHeight
    mFitOn = True
    Err.Clear
    On Error GoTo Fail

    rng.Columns.AutoFit
    rng.Rows.AutoFit

    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    Application.OnUndo "Undo AutoFit Rows / Columns", MacroPrefix() & "UndoAutoFit"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "Row and column sizes fitted. (" & rng.Address(False, False) & ")" & _
                            "  |  Ctrl+Z to undo"
    ScheduleStatusClear
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "AutoFit failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "AutoFit Rows and Columns"
End Sub

'------------------------------------------------------------
' 6. ExportSheetAsText (Ctrl+Shift+E)
'    pull the current sheet out into a new workbook, then
'    open the save dialog as [sourcename_yyyymmdd.txt] / Text (tab delimited).
'------------------------------------------------------------
Public Sub ExportSheetAsText()
    Dim srcWb As Workbook
    Dim sh As Object
    Dim newWb As Workbook
    Dim sBase As String, sDefault As String
    Dim vPath As Variant
    Dim oldAlerts As Boolean
    Dim step2 As String
    Dim errNo As Long, errDesc As String
    Dim saveErr As Long, saveDesc As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "There is no sheet to save.", vbExclamation, "Export Sheet as Text"
        Exit Sub
    End If

    oldAlerts = Application.DisplayAlerts       ' for restoring on error. must be captured first

    Set srcWb = ActiveWorkbook
    Set sh = ActiveSheet

    ' capture the source name before copying (ActiveWorkbook changes afterwards)
    If srcWb.Path = "" Then
        sBase = sh.Name
    Else
        sBase = srcWb.Name
        If InStrRev(sBase, ".") > 1 Then sBase = Left$(sBase, InStrRev(sBase, ".") - 1)
    End If
    sDefault = SafeFileName(sBase) & "_" & Format$(Date, "yyyymmdd") & ".txt"

    On Error GoTo Fail

    step2 = "Copy sheet to a new workbook"
    sh.Copy
    Set newWb = ActiveWorkbook
    If newWb Is Nothing Then
        MsgBox "Could not create a new window.", vbExclamation, "Export Sheet as Text"
        Exit Sub
    End If

    step2 = "Choose where to save"
    vPath = Application.GetSaveAsFilename( _
                InitialFileName:=sDefault, _
                FileFilter:="Text (tab delimited) (*.txt),*.txt", _
                Title:="Save As")

    If VarType(vPath) = vbBoolean Then
        Exit Sub                                ' cancelled - leave the copied workbook as it is
    End If

    step2 = "Save file"
    Application.DisplayAlerts = False

    ' first try : Text (tab delimited) - Windows
    On Error Resume Next
    newWb.SaveAs Filename:=CStr(vPath), FileFormat:=xlTextWindows
    saveErr = Err.Number
    Err.Clear
    On Error GoTo Fail

    ' second try : the current platform's text format
    If saveErr <> 0 Then
        On Error Resume Next
        newWb.SaveAs Filename:=CStr(vPath), FileFormat:=xlCurrentPlatformText
        saveErr = Err.Number
        saveDesc = Err.Description
        Err.Clear
        On Error GoTo Fail
    End If

    Application.DisplayAlerts = oldAlerts

    If saveErr <> 0 Then
        MsgBox "Could not save." & vbCrLf & vbCrLf & _
               "Saved to: " & CStr(vPath) & vbCrLf & _
               "Error " & saveErr & ": " & saveDesc & vbCrLf & vbCrLf & _
               "The copied window was left open." & vbCrLf & _
               "Try pressing F12 and saving it yourself.", vbExclamation, "Export Sheet as Text"
        Exit Sub
    End If

    MsgBox "Saved." & vbCrLf & vbCrLf & newWb.FullName, _
           vbInformation, "Export Sheet as Text"
    Exit Sub

Fail:
    errNo = Err.Number                          ' must be captured before Err.Clear
    errDesc = Err.Description

    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    Err.Clear
    On Error GoTo 0

    MsgBox "Could not save." & vbCrLf & vbCrLf & _
           "Step: " & step2 & vbCrLf & _
           "Error " & errNo & ": " & errDesc & vbCrLf & _
           IIf(Len(CStr(vPath)) > 0 And VarType(vPath) <> vbBoolean, _
               "Saved to: " & CStr(vPath) & vbCrLf, "") & vbCrLf & _
           "The new window was left open so you can save it yourself.", _
           vbExclamation, "Export Sheet as Text"
End Sub

Private Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant
    Dim i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next i
    s = Trim$(s)
    If Len(s) = 0 Then s = "sheet"
    SafeFileName = s
End Function

'------------------------------------------------------------
' 7. HighlightDuplicateRows (Ctrl+Shift+D)
'------------------------------------------------------------
Public Sub HighlightDuplicateRows()
    Dim rng As Range
    Dim dic As Object
    Dim arr As Variant
    Dim key As String
    Dim r As Long, c As Long
    Dim rowCnt As Long, colCnt As Long
    Dim dupRows As Long
    Dim oldScreen As Boolean
    Dim groupColor As Object
    Dim groupCount As Long
    Dim palette As Variant
    Dim targetRows() As Variant
    Dim prevColor() As Variant
    Dim prevColorIndex() As Variant
    Dim cellColor() As Variant
    Dim cellColorIndex() As Variant
    Dim visible() As Boolean
    Dim visibleRange As Range
    Dim area As Range
    Dim startIdx As Long
    Dim visibleCount As Long
    Dim colorAddr() As String
    Dim colorNo As Long
    Dim i As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell first.", vbExclamation, "Highlight Duplicate Rows"
        Exit Sub
    End If

    If Selection.Cells.CountLarge >= 2 Then
        Set rng = Selection
    Else
        Set rng = SkipNonHeaderRows(ActiveCell.CurrentRegion)
    End If

    ' guard against a whole row or column being selected
    On Error Resume Next
    Set rng = Application.Intersect(rng, rng.Worksheet.UsedRange)
    Err.Clear
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "There is no data to check.", vbExclamation, "Highlight Duplicate Rows"
        Exit Sub
    End If
    ' when several spots are picked with Ctrl, compare cell by cell
    If rng.Areas.Count > 1 Then
        FindDuplicateValues rng
        Exit Sub
    End If

    rowCnt = rng.Rows.Count
    colCnt = rng.Columns.Count

    If rowCnt < 2 Then
        MsgBox "Please select two or more rows.", vbExclamation, "Highlight Duplicate Rows"
        Exit Sub
    End If
    '--- work out the visible rows (filtered and hidden rows left out) ---
    ReDim visible(1 To rowCnt)
    visibleCount = 0

    On Error Resume Next
    Set visibleRange = rng.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo 0

    If visibleRange Is Nothing Then
        For r = 1 To rowCnt
            visible(r) = True
        Next r
        visibleCount = rowCnt
    Else
        For Each area In visibleRange.Areas
            startIdx = area.Row - rng.Row + 1
            For r = startIdx To startIdx + area.Rows.Count - 1
                If r >= 1 And r <= rowCnt Then
                    If Not visible(r) Then
                        visible(r) = True
                        visibleCount = visibleCount + 1
                    End If
                End If
            Next r
        Next area
    End If

    If visibleCount < 2 Then
        MsgBox "Fewer than two visible rows." & vbCrLf & _
               "Please check the filter, then run again.", vbExclamation, "Highlight Duplicate Rows"
        Exit Sub
    End If

    '--- report against the number of rows actually checked ---
    If visibleCount > 100000 Then
        If MsgBox(Format$(visibleCount, "#,##0") & " rows will be checked. This may take a while." & vbCrLf & _
                  IIf(visibleCount < rowCnt, _
                      "(total " & Format$(rowCnt, "#,##0") & " rows, visible ones only)" & vbCrLf, "") & _
                  "Do you want to continue?", vbOKCancel + vbQuestion, "Highlight Duplicate Rows") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating

    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = rng.Value

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set groupColor = CreateObject("Scripting.Dictionary")
    groupColor.CompareMode = 1

    palette = ColorPalette()

    ' pass 1 : count by key (visible rows only)
    For r = 1 To rowCnt
      If visible(r) Then
        key = ""
        For c = 1 To colCnt
            key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic.Exists(key) Then
            dic(key) = dic(key) + 1
        Else
            dic.Add key, 1
        End If
      End If
    Next r

    ' pass 2 : collect the duplicate row numbers (visible rows only)
    ReDim targetRows(1 To rowCnt)
    For r = 1 To rowCnt
      If visible(r) Then
        key = ""
        For c = 1 To colCnt
            key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            dupRows = dupRows + 1
            targetRows(dupRows) = r
        End If
      End If
    Next r

    If dupRows > 0 Then
        ' keep the existing colors so it can be undone
        ' read each row's color in one go. when a row holds mixed colors (Null)
        ' that row alone is kept cell by cell.
        ReDim prevColor(1 To dupRows)
        ReDim prevColorIndex(1 To dupRows)
        For r = 1 To dupRows
            With rng.Rows(targetRows(r))
                prevColor(r) = .Interior.Color
                prevColorIndex(r) = .Interior.ColorIndex
            End With
            If IsNull(prevColor(r)) Or IsNull(prevColorIndex(r)) Then
                ReDim cellColor(1 To colCnt)
                ReDim cellColorIndex(1 To colCnt)
                For c = 1 To colCnt
                    cellColor(c) = rng.Cells(targetRows(r), c).Interior.Color
                    cellColorIndex(c) = rng.Cells(targetRows(r), c).Interior.ColorIndex
                Next c
                prevColor(r) = cellColor
                prevColorIndex(r) = cellColorIndex
            End If
        Next r
        mDupWb = rng.Worksheet.Parent.Name
        mDupWs = rng.Worksheet.Name
        mDupAddr = rng.Address
        mDupRows = targetRows
        mDupColor = prevColor
        mDupIndex = prevColorIndex
        mDupCount = dupRows
    End If

    ' pass 3 : gather addresses per color and paint in one go (painting row by row is very slow)
    dupRows = 0
    ReDim colorAddr(0 To UBound(palette))
    For r = 1 To rowCnt
      If visible(r) Then
        key = ""
        For c = 1 To colCnt
            key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            If Not groupColor.Exists(key) Then
                groupColor.Add key, groupCount Mod (UBound(palette) + 1)
                groupCount = groupCount + 1
            End If
            colorNo = groupColor(key)
            If Len(colorAddr(colorNo)) > 0 Then colorAddr(colorNo) = colorAddr(colorNo) & ","
            colorAddr(colorNo) = colorAddr(colorNo) & rng.Rows(r).Address(False, False)
            dupRows = dupRows + 1

            ' when the address string gets long, paint and clear it there and then
            If Len(colorAddr(colorNo)) > 200 Then
                rng.Worksheet.Range(colorAddr(colorNo)).Interior.Color = palette(colorNo)
                colorAddr(colorNo) = ""
            End If
        End If
      End If
    Next r

    For i = 0 To UBound(palette)
        If Len(colorAddr(i)) > 0 Then
            rng.Worksheet.Range(colorAddr(i)).Interior.Color = palette(i)
        End If
    Next i

    Application.ScreenUpdating = oldScreen

    If dupRows > 0 Then
        On Error Resume Next
        Application.OnUndo "Undo Highlight Duplicate Rows", MacroPrefix() & "UndoDuplicateColors"
        Err.Clear
        On Error GoTo Fail
    End If

    If dupRows = 0 Then
        MsgBox "No duplicate rows found." & vbCrLf & _
               "Checked range: " & rng.Address(False, False) & _
               IIf(visibleCount < rowCnt, "  (visible " & visibleCount & " rows checked)", ""), _
               vbInformation, "Highlight Duplicate Rows"
    Else
        MsgBox "Duplicates " & groupCount & " kinds / " & Format$(dupRows, "#,##0") & " rows highlighted." & vbCrLf & _
               "Matching rows share the same color." & vbCrLf & _
               "Checked range: " & rng.Address(False, False) & _
               IIf(visibleCount < rowCnt, "  (visible " & visibleCount & " rows checked)", "") & _
               IIf(groupCount > UBound(palette) + 1, vbCrLf & "There are " & (UBound(palette) + 1) & _
                   " kinds, so some colors repeat.", "") & vbCrLf & vbCrLf & _
               "You can undo this with Ctrl+Z." & vbCrLf & _
               "To clear only the colors : select the range, then Ctrl+Shift+C", _
               vbInformation, "Highlight Duplicate Rows"
    End If
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "An error occurred while checking for duplicates." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Highlight Duplicate Rows"
End Sub

Public Sub ClearFillColor()
    Dim rng As Range
    Dim visibleRange As Range
    Dim target As Range
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the range whose fill color should be cleared, then run again.", vbExclamation, "Clear Cell Fill"
        Exit Sub
    End If

    Set rng = Selection

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    '--- visible cells only (filtered and hidden rows left out) ---
    On Error Resume Next
    Set visibleRange = rng.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo Fail

    If visibleRange Is Nothing Then
        Set target = rng
    Else
        Set target = visibleRange
    End If

    ' kept so it can be undone
    SaveFillColors target

    target.Interior.ColorIndex = xlColorIndexNone

    Application.ScreenUpdating = oldScreen

    If mBgOn Then
        On Error Resume Next
        Application.OnUndo "Undo Clear Cell Fill", MacroPrefix() & "UndoFillColors"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "Cell fill color cleared. (" & target.Address(False, False) & ")" & _
                            IIf(mBgOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not clear the fill color." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Clear Cell Fill"
End Sub

'------------------------------------------------------------
' 8. CopyFilePath (Ctrl+Shift+N)
'------------------------------------------------------------
Public Sub CopyFilePath()
    Dim s As String

    If ActiveWorkbook Is Nothing Then Exit Sub

    If ActiveWorkbook.Path = "" Then
        MsgBox "This file has not been saved yet.", vbExclamation, "Copy File Path"
        Exit Sub
    End If

    s = ActiveWorkbook.FullName

    If ToClipboard(s) Then
        MsgBox "The file path was copied." & vbCrLf & vbCrLf & s, _
               vbInformation, "Copy File Path"
    Else
        InputBox "Copying failed. Please copy the path below yourself (Ctrl+C).", _
                 "Copy File Path", s
    End If
End Sub

Private Function ToClipboard(ByVal s As String) As Boolean
    Dim obj As Object
    On Error GoTo Fail
    ' create MSForms.DataObject without adding a reference (late binding)
    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    obj.SetText s
    obj.PutInClipboard
    ToClipboard = True
    Exit Function
Fail:
    ToClipboard = False
End Function


'============================================================
'  F. Work tools menu (Ctrl+Shift+M)
'     a CommandBar popup instead of a UserForm - no reference needed
'============================================================

Public Sub MacroLauncher()
    Dim cb As CommandBar

    On Error Resume Next
    Application.CommandBars(BAR_MAIN_MENU).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=BAR_MAIN_MENU, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, _
                                         Temporary:=True)

    AddMenuItem cb, "Open New Window (same file in two windows)", "OpenNewWindow", "Ctrl+Shift+O"
    AddMenuItem cb, "Copy to New Workbook", "CopyToNewWorkbook", "Ctrl+Shift+Q"
    AddMenuItem cb, "Comma Format", "CommaFormat", "Ctrl+Shift+X"
    AddMenuItem cb, "Text to Columns - as General", "TextToGeneral", "Ctrl+Shift+Y"
    AddMenuItem cb, "Text to Columns - as Text", "TextToText", "Ctrl+Shift+I"
    AddMenuItem cb, "Table Borders", "TableBorders", "Ctrl+Shift+U"
    AddMenuItem cb, "Center Align / Default", "ToggleCenterAlign", "Ctrl+Alt+Z"
    AddMenuItem cb, "Format Brush", "FormatBrush", "Ctrl+Alt+X"
    AddMenuItem cb, "Center Across Selection (no merging)", "CenterAcrossSelection", "Ctrl+Alt+S"
    AddMenuItem cb, "Flash Fill (upgraded Ctrl+E)", "FlashFillToEnd", "Ctrl+Alt+A"
    AddMenuItem cb, "Auto Fill Series (1,2,3 / Mon,Tue,Wed)", "AutoFillSeries", "Ctrl+Alt+Q"
    AddMenuItem cb, "Fill Blanks From Above", "FillBlanksFromAbove", "Ctrl+Shift+4"
    AddMenuItem cb, "Fill Formula Down", "FillFormulaDown", "Ctrl+Shift+1"
    AddMenuItem cb, "Insert Index Column (No)", "InsertIndexColumn", "Ctrl+Shift+2"
    AddMenuItem cb, "Select to End of Table", "SelectToEnd", "Ctrl+Shift+3"
    AddMenuItem cb, "Remove Rows Above Header", "RemoveRowsAboveHeader", "Ctrl+Alt+E"
    AddMenuItem cb, "AutoFit Rows / Columns", "AutoFitRowsCols", "Ctrl+Shift+W"
    AddMenuItem cb, "Freeze / unfreeze header row", "FreezeHeaderRow", "Ctrl+Shift+S"
    AddMenuItem cb, "Tidy All Sheets (every sheet to A1)", "TidyAllSheets", "Ctrl+Shift+H"

    AddMenuItem cb, "Create PivotTable (new sheet)", "CreatePivot", "Ctrl+Shift+J", True
    AddMenuItem cb, "Classic PivotTable Layout (cannot be undone)", "ClassicPivotCleanup", "Ctrl+Shift+T"
    AddMenuItem cb, "Highlight Duplicate Rows", "HighlightDuplicateRows", "Ctrl+Shift+D"
    AddMenuItem cb, "Find Duplicates by Column (highlight · number · extract)", "ColumnDuplicates", "Ctrl+Shift+V"
    AddMenuItem cb, "Compare Two Columns (between lists · within each row)", "CompareTwoColumns", "Ctrl+Alt+D"
    AddMenuItem cb, "Multi VLOOKUP (group and spread out)", "MultiLookup", "Ctrl+Shift+L"
    AddMenuItem cb, "Column Headers to Row Data (unpivot)", "Unpivot", "Ctrl+Alt+B"
    AddMenuItem cb, "Clear Cell Fill (select a range first)", "ClearFillColor", "Ctrl+Shift+C"

    AddMenuItem cb, "Turn filter on / off", "ToggleAutoFilter", "Ctrl+Shift+A", True
    AddMenuItem cb, "Filter by Selected Cells", "FilterBySelectedValue", "Ctrl+Shift+Z"
    AddMenuItem cb, "Run Advanced Filter", "AdvancedFilter", "Ctrl+Shift+F"
    AddMenuItem cb, "Set Up Advanced Filter Criteria", "BuildCriteriaBlock", "Ctrl+Shift+R"
    AddMenuItem cb, "Advanced Filter Range", "ShowAdvancedFilterRange", ""
    AddMenuItem cb, "SUBTOTAL Sum", "SubtotalSum", "Ctrl+Shift+G"
    AddMenuItem cb, "SUBTOTAL Count (filled cells)", "SubtotalCount", "Ctrl+Alt+G"

    AddMenuItem cb, "Export Sheet as Text (for txt->dat conversion)", "ExportSheetAsText", "Ctrl+Shift+E", True
    AddMenuItem cb, "Copy File Path", "CopyFilePath", "Ctrl+Shift+N"
    AddMenuItem cb, "Find by Value (matches the result)", "FindByValue", "Ctrl+Alt+F"

    AddMenuItem cb, "Set Password", "SetPassword", "Ctrl+Shift+P", True
    AddMenuItem cb, AskOnCloseCaption(), "ToggleAskOnClose", ""
    AddMenuItem cb, "Change default password (now : " & ReadDefaultPassword() & ")", "ChangeDefaultPassword", ""

    AddMenuItem cb, "Paste / remove the shortcut list", "CheatSheet", "Ctrl+Shift+K", True
    AddMenuItem cb, "Comma format look :  " & CommaFormatOneLine(), "CommaFormatMenu", ""

    AddMenuItem cb, "Reset Settings (run before distributing)", "ResetSettings", ""
    AddMenuItem cb, "Restart Work Tools", "WorkTools_Restart", ""
    AddMenuItem cb, "Diagnostics / Version", "WorkTools_Diagnostics", ""

    cb.ShowPopup
    Exit Sub

Fail:
    MsgBox "Could not show the menu." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Work Tools"
End Sub

Private Sub AddMenuItem(ByVal cb As CommandBar, _
                     ByVal shown As String, _
                     ByVal macroName As String, _
                     ByVal keyLabel As String, _
                     Optional ByVal separator As Boolean = False)
    Dim btn As CommandBarButton

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    If Len(keyLabel) > 0 Then
        btn.Caption = shown & "   (" & keyLabel & ")"
    Else
        btn.Caption = shown
    End If
    btn.OnAction = MacroPrefix() & macroName
    btn.BeginGroup = separator
    btn.Style = msoButtonCaption
End Sub

Private Sub AddMenuTitle(ByVal cb As CommandBar, ByVal shown As String)
    Dim btn As CommandBarButton
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = shown
    btn.Style = msoButtonCaption
    btn.Enabled = False
End Sub

Private Function AskOnCloseCaption() As String
    If AskOnCloseEnabled() Then
        AskOnCloseCaption = "Ask for a password on close : on  ->  turn off"
    Else
        AskOnCloseCaption = "Ask for a password on close : off  ->  turn on"
    End If
End Function


'============================================================
'  G. Advanced filter
'     the criteria block sits below the data block with a blank line between.
'     both blocks are found on their own, so no range has to be given each time.
'============================================================

'------------------------------------------------------------
' find the first data block on the sheet
'------------------------------------------------------------
Private Function DataBlock(ByVal ws As Worksheet) As Range
    Dim ur As Range
    Dim r As Long, c As Long
    Dim endRow As Long, endCol As Long

    On Error GoTo Fail
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function

    endRow = ur.Row + ur.Rows.Count - 1
    endCol = ur.Column + ur.Columns.Count - 1

    For r = ur.Row To endRow
        If Application.CountA(ws.Rows(r)) > 0 Then
            For c = ur.Column To endCol
                If Len(CStr(ws.Cells(r, c).Value)) > 0 Then
                    Set DataBlock = SkipNonHeaderRows(ws.Cells(r, c).CurrentRegion)
                    Exit Function
                End If
            Next c
        End If
    Next r
    Exit Function
Fail:
End Function

'------------------------------------------------------------
' when the block's first row is a SUBTOTAL row, the row below is taken as the header
'------------------------------------------------------------
Private Function SkipNonHeaderRows(ByVal rg As Range) As Range
    Dim skipped As Long
    Dim leftRows As Long

    Set SkipNonHeaderRows = rg
    If rg Is Nothing Then Exit Function
    If rg.Rows.Count < 3 Then Exit Function

    On Error GoTo Done

    ' strip up to three non-header rows off the top
    skipped = 0
    Do While skipped < 3
        leftRows = rg.Rows.Count - skipped
        If leftRows < 3 Then Exit Do
        If Not NotAHeaderRow(rg, skipped + 1) Then Exit Do
        skipped = skipped + 1
    Loop

    If skipped > 0 Then
        Set SkipNonHeaderRows = rg.Offset(skipped, 0).Resize(rg.Rows.Count - skipped, rg.Columns.Count)
    End If

Done:
End Function

'------------------------------------------------------------
' decide whether that row is hard to read as a header
'   1) it holds a SUBTOTAL formula
'   2) most filled cells are numbers  (headers are usually text)
'   3) far fewer filled cells than the row below
'------------------------------------------------------------
Private Function NotAHeaderRow(ByVal rg As Range, ByVal rowNo As Long) As Boolean
    Dim i As Long
    Dim filled As Long, numCount As Long
    Dim belowFilled As Long
    Dim v As Variant
    Dim f As String
    Dim nC As Long

    On Error GoTo Done
    nC = rg.Columns.Count
    If rowNo + 1 > rg.Rows.Count Then Exit Function

    For i = 1 To nC
        ' 1) a SUBTOTAL formula
        If rg.Cells(rowNo, i).HasFormula Then
            f = UCase$(CStr(rg.Cells(rowNo, i).Formula))
            If InStr(f, "SUBTOTAL") > 0 Then
                NotAHeaderRow = True
                Exit Function
            End If
        End If

        v = rg.Cells(rowNo, i).Value
        If Len(CStr(v)) > 0 Then
            filled = filled + 1
            ' only count it when the cell really is stored as a number.
            '   IsDate must not be used here - it reads header text like "Jan 2025"
            '   as a date and ends up stripping off a real header row.
            '   dates are not counted either - monthly data sometimes has real date
            '   cells as headers, and subtotal rows hold numbers, not dates.
            Select Case VarType(v)
                Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
                    numCount = numCount + 1
            End Select
        End If

        If Len(CStr(rg.Cells(rowNo + 1, i).Value)) > 0 Then belowFilled = belowFilled + 1
    Next i

    If filled = 0 Then
        NotAHeaderRow = True                      ' a completely empty row
        Exit Function
    End If

    ' 2) if over half the filled cells are numbers or dates, it is hard to read as a header
    If numCount * 2 >= filled Then
        NotAHeaderRow = True
        Exit Function
    End If

    ' 3) if fewer than half as many cells are filled as the row below, it is not a header
    If belowFilled >= 4 Then
        If filled * 2 < belowFilled Then NotAHeaderRow = True
    End If

Done:
End Function
'------------------------------------------------------------
' find the criteria block below the data block (Nothing when there is none)
'------------------------------------------------------------
Private Function CriteriaBlock(ByVal ws As Worksheet, ByVal data As Range) As Range
    Dim ur As Range
    Dim r As Long, c As Long
    Dim startIdx As Long, endRow As Long, endCol As Long

    On Error GoTo Fail
    If data Is Nothing Then Exit Function

    Set ur = ws.UsedRange
    startIdx = data.Row + data.Rows.Count
    endRow = ur.Row + ur.Rows.Count - 1
    endCol = ur.Column + ur.Columns.Count - 1
    If startIdx > endRow Then Exit Function

    For r = startIdx To endRow
        If Application.CountA(ws.Rows(r)) > 0 Then
            For c = ur.Column To endCol
                If Len(CStr(ws.Cells(r, c).Value)) > 0 Then
                    Set CriteriaBlock = ws.Cells(r, c).CurrentRegion
                    Exit Function
                End If
            Next c
        End If
    Next r
    Exit Function
Fail:
End Function

'------------------------------------------------------------
' check that every criteria header exists among the data headers
'------------------------------------------------------------
Private Function CheckHeaders(ByVal data As Range, ByVal criteria As Range, _
                            ByRef badName As String) As Boolean
    Dim i As Long, j As Long
    Dim h As String
    Dim found As Boolean

    badName = ""
    On Error GoTo Fail

    For i = 1 To criteria.Columns.Count
        h = Trim$(CStr(criteria.Cells(1, i).Value))
        If Len(h) > 0 Then
            found = False
            For j = 1 To data.Columns.Count
                If StrComp(h, Trim$(CStr(data.Cells(1, j).Value)), vbTextCompare) = 0 Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found Then
                badName = h
                Exit Function
            End If
        End If
    Next i

    CheckHeaders = True
    Exit Function
Fail:
    CheckHeaders = True                       ' when it cannot be decided, let it through and leave it to Excel's own error
End Function

'------------------------------------------------------------
' advanced filter (Ctrl+Shift+F)
'   builds the criteria block when there is none, and runs the filter when there is.
'------------------------------------------------------------
Public Sub AdvancedFilter()
    Dim ws As Worksheet
    Dim data As Range, criteria As Range
    Dim badName As String
    Dim n As Long

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "AdvancedFilter"
        Exit Sub
    End If
    Set ws = ActiveSheet

    '--- clear the advanced filter when one is already on (toggle) ---
    On Error Resume Next
    If ws.FilterMode And Not ws.AutoFilterMode Then
        ws.ShowAllData
        Err.Clear
        On Error GoTo 0
        Application.StatusBar = "Advanced Filter cleared. Press again to filter by the criteria."
        ScheduleStatusClear
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    Set data = DataBlock(ws)
    If data Is Nothing Then
        MsgBox "Could not find any data.", vbExclamation, "AdvancedFilter"
        Exit Sub
    End If
    If data.Rows.Count < 2 Then
        MsgBox "A header row and data rows are both required.", vbExclamation, "AdvancedFilter"
        Exit Sub
    End If

    Set criteria = CriteriaBlock(ws, data)

    '--- build the criteria block when there is none ---
    If criteria Is Nothing Then
        BuildCriteriaBlock
        Exit Sub
    End If

    '--- when no criteria values are typed yet, jump to where they go ---
    If criteria.Rows.Count < 2 Then
        On Error Resume Next
        Application.Goto ws.Cells(criteria.Row + 1, criteria.Column), False
        ws.Cells(criteria.Row + 1, criteria.Column).Select
        Err.Clear
        On Error GoTo 0
        MsgBox "The criteria block has no values yet." & vbCrLf & _
               "Type the criteria under the headers, then run again.", _
               vbExclamation, "AdvancedFilter"
        Exit Sub
    End If

    '--- header check ---
    If Not CheckHeaders(data, criteria, badName) Then
        MsgBox "Criteria header '" & badName & "' is not in the data headers." & vbCrLf & _
               "Please make the name match the data exactly.", vbExclamation, "AdvancedFilter"
        Exit Sub
    End If

    On Error GoTo Fail

    ' clear the previous filter state first
    On Error Resume Next
    ws.ShowAllData
    Err.Clear
    On Error GoTo Fail

    data.AdvancedFilter Action:=xlFilterInPlace, _
                          CriteriaRange:=criteria, _
                          Unique:=False

    n = VisibleRowCount(data)

    Application.StatusBar = "Advanced Filter applied  |  list " & data.Address(False, False) & _
                            "  Criteria " & criteria.Address(False, False) & _
                            "  Result " & Format$(n, "#,##0") & " rows  |  clear with Ctrl+Shift+R"
    ScheduleStatusClear
    Exit Sub

Fail:
    MsgBox "Could not apply the Advanced Filter." & vbCrLf & vbCrLf & _
           "List range: " & data.Address(False, False) & vbCrLf & _
           "Criteria range: " & criteria.Address(False, False) & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "AdvancedFilter"
End Sub

'------------------------------------------------------------
' how many rows the filter left (header not counted)
'------------------------------------------------------------
Private Function VisibleRowCount(ByVal data As Range) As Long
    Dim rng As Range
    On Error GoTo Fail
    Set rng = data.Columns(1).Offset(1, 0).Resize(data.Rows.Count - 1, 1)
    VisibleRowCount = rng.SpecialCells(xlCellTypeVisible).Cells.Count
    Exit Function
Fail:
    VisibleRowCount = 0
End Function

'------------------------------------------------------------
' build the criteria block - copy the data headers below
'------------------------------------------------------------
Public Sub BuildCriteriaBlock()
    Dim errNo As Long, errDesc As String
    Dim ws As Worksheet
    Dim data As Range, existing As Range
    Dim header As Range, target As Range
    Dim undoRange As Range
    Dim pasteRow As Long
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Build Criteria Block"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set data = DataBlock(ws)
    If data Is Nothing Then
        MsgBox "Could not find any data.", vbExclamation, "Build Criteria Block"
        Exit Sub
    End If

    Set existing = CriteriaBlock(ws, data)
    If Not existing Is Nothing Then
        If MsgBox("Criteria columns already exist. (" & existing.Address(False, False) & ")" & vbCrLf & _
                  "Clear it and build a new one?" & vbCrLf & vbCrLf & _
                  "You can undo this with Ctrl+Z.", _
                  vbYesNo + vbQuestion, "Set Up Criteria Columns") = vbNo Then Exit Sub
    End If

    On Error GoTo Fail

    Set header = data.Rows(1)
    pasteRow = data.Row + data.Rows.Count      ' end of data + 1 = the blank line
    pasteRow = pasteRow + 1                          ' the criteria headers go below that

    Set target = ws.Cells(pasteRow, data.Column)

    '--- kept so it can be undone : before clearing, both the old block and where the new one goes ---
    r1 = pasteRow
    r2 = pasteRow
    c1 = data.Column
    c2 = data.Column + data.Columns.Count - 1

    If Not existing Is Nothing Then
        If existing.Row < r1 Then r1 = existing.Row
        If existing.Row + existing.Rows.Count - 1 > r2 Then r2 = existing.Row + existing.Rows.Count - 1
        If existing.Column < c1 Then c1 = existing.Column
        If existing.Column + existing.Columns.Count - 1 > c2 Then c2 = existing.Column + existing.Columns.Count - 1
    End If

    Set undoRange = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    SaveCriteriaBlock undoRange

    '--- now clear it ---
    If Not existing Is Nothing Then
        On Error Resume Next
        existing.Clear
        Err.Clear
        On Error GoTo Fail
    End If

    header.Copy
    target.PasteSpecial xlPasteValues
    target.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    On Error Resume Next
    Application.OnUndo "Undo Set Up Criteria Columns", MacroPrefix() & "UndoCriteriaBlock"
    Err.Clear
    On Error GoTo Fail

    On Error Resume Next
    ws.Cells(pasteRow + 1, data.Column).Select
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "Criteria columns are ready. Type the criteria, then press Ctrl+Shift+F." & _
                            "  |  Ctrl+Z to undo"
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    Err.Clear
    On Error GoTo 0
    MsgBox "Could not set up the criteria columns." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Set Up Criteria Columns"
End Sub


'------------------------------------------------------------
' check the advanced filter ranges - shows what was read as data and as criteria
'------------------------------------------------------------
Public Sub ShowAdvancedFilterRange()
    Dim ws As Worksheet
    Dim data As Range, criteria As Range
    Dim s As String
    Dim blankLine As Long
    Dim badName As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Advanced Filter Range"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set data = DataBlock(ws)

    If data Is Nothing Then
        MsgBox "Could not find the data block." & vbCrLf & _
               "Please check whether the sheet is empty.", vbExclamation, "Advanced Filter Range"
        Exit Sub
    End If

    s = "■ Data block" & vbCrLf & _
        "   Range : " & data.Address(False, False) & vbCrLf & _
        "   Size : " & Format$(data.Rows.Count, "#,##0") & " rows x " & _
                       data.Columns.Count & "Column" & vbCrLf & _
        "   Header : " & HeaderList(data) & vbCrLf & vbCrLf

    Set criteria = CriteriaBlock(ws, data)

    If criteria Is Nothing Then
        blankLine = data.Row + data.Rows.Count
        s = s & "■ Criteria block  ->  none" & vbCrLf & vbCrLf & _
                "Data " & (data.Row + data.Rows.Count - 1) & " and ends at row " & vbCrLf & _
                "There are no rows with values below it." & vbCrLf & vbCrLf & _
                "Running [Build a new criteria block] will" & vbCrLf & _
                blankLine & " rows will be cleared and " & (blankLine + 1) & " the headers go into row "
        MsgBox s, vbInformation, "Advanced Filter Range"
        Exit Sub
    End If

    s = s & "■ Criteria block" & vbCrLf & _
            "   Range : " & criteria.Address(False, False) & vbCrLf & _
            "   Size : " & criteria.Rows.Count & " rows x " & criteria.Columns.Count & "Column" & vbCrLf & _
            "   Header : " & HeaderList(criteria) & vbCrLf & vbCrLf

    If criteria.Rows.Count < 2 Then
        s = s & "* No criteria values yet." & vbCrLf & _
                "   Below the header : " & (criteria.Row + 1) & "Type the criteria into row "
    ElseIf Not CheckHeaders(data, criteria, badName) Then
        s = s & "* Criteria header '" & badName & "'" & vbCrLf & _
                " is not in the data headers. Please check the spelling."
    Else
        s = s & "○ Ready to filter.  (Ctrl+Shift+F)"
    End If

    MsgBox s, vbInformation, "Advanced Filter Range"
End Sub

Private Function HeaderList(ByVal rng As Range) As String
    Dim i As Long
    Dim s As String
    Dim v As String

    On Error GoTo Fail
    For i = 1 To rng.Columns.Count
        v = Trim$(CStr(rng.Cells(1, i).Value))
        If Len(v) = 0 Then v = "(blank)"
        If Len(s) > 0 Then s = s & " | "
        s = s & v
        If i >= 8 Then
            s = s & " | ..."
            Exit For
        End If
    Next i
    HeaderList = s
    Exit Function
Fail:
    HeaderList = "(check failed)"
End Function

'------------------------------------------------------------
' turn AutoFilter on / off (Ctrl+Shift+A)
'   inside a table (ListObject), the table's own filter buttons are toggled.
'------------------------------------------------------------
Public Sub ToggleAutoFilter()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lo As Object
    Dim autoBtn As Boolean
    Dim isOn As Boolean

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Filter"
        Exit Sub
    End If
    Set ws = ActiveSheet

    On Error GoTo Fail

    '--- inside a table (ListObject), toggle the table's filter ---
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    Err.Clear
    On Error GoTo Fail

    If Not lo Is Nothing Then
        lo.ShowAutoFilter = Not lo.ShowAutoFilter
        Application.StatusBar = "Table '" & lo.Name & "' filter " & _
                                IIf(lo.ShowAutoFilter, " is now on.", " turned off.")
        ScheduleStatusClear
        Exit Sub
    End If

    On Error Resume Next
    autoBtn = ws.AutoFilterMode
    isOn = ws.FilterMode
    Err.Clear
    On Error GoTo Fail

    '--- when rows are hidden by an advanced filter, clear that first ---
    If isOn And Not autoBtn Then
        On Error Resume Next
        ws.ShowAllData
        Err.Clear
        On Error GoTo Fail
        Application.StatusBar = "Advanced Filter cleared. All rows are shown."
        ScheduleStatusClear
        Exit Sub
    End If

    '--- turn AutoFilter off when it is on (same as the ribbon's [Filter] button) ---
    If autoBtn Then
        On Error Resume Next
        If isOn Then ws.ShowAllData
        Err.Clear
        On Error GoTo Fail
        ws.AutoFilterMode = False
        Application.StatusBar = "The filter was turned off."
        ScheduleStatusClear
        Exit Sub
    End If

    '--- turn it on when it is not ---
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.CountLarge >= 2 Then
            Set rng = SkipNonHeaderRows(Selection)
        Else
            Set rng = SkipNonHeaderRows(ActiveCell.CurrentRegion)
        End If
    End If

    If rng Is Nothing Then
        MsgBox "Could not find a range to filter.", vbExclamation, "Filter"
        Exit Sub
    End If
    If rng.Cells.CountLarge < 2 Then
        MsgBox "Put the cursor on a cell with data and run again.", vbExclamation, "Filter"
        Exit Sub
    End If

    rng.AutoFilter

    Application.StatusBar = "The filter was turned on. (" & rng.Address(False, False) & ")"
    ScheduleStatusClear
    Exit Sub

Fail:
    MsgBox "Could not change the filter." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Filter"
End Sub

'------------------------------------------------------------
' freeze / unfreeze the header row (Ctrl+Shift+S)
'   finds the header row and freezes below it. press again to unfreeze.
'------------------------------------------------------------
Public Sub FreezeHeaderRow()
    Dim ws As Worksheet
    Dim data As Range
    Dim headerRow As Long
    Dim undoCell As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Freeze Header Row"
        Exit Sub
    End If
    Set ws = ActiveSheet

    On Error GoTo Fail

    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
        Application.StatusBar = "Freeze panes was turned off."
        ScheduleStatusClear
        Exit Sub
    End If

    '--- find the header row ---
    headerRow = 1
    Set data = DataBlock(ws)
    If Not data Is Nothing Then headerRow = data.Row

    If TypeName(Selection) = "Range" Then Set undoCell = Selection

    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Err.Clear
    On Error GoTo Fail

    ws.Cells(headerRow + 1, 1).Select
    ActiveWindow.FreezePanes = True

    On Error Resume Next
    If Not undoCell Is Nothing Then undoCell.Select
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = headerRow & " is frozen. Press again to unfreeze."
    ScheduleStatusClear
    Exit Sub

Fail:
    MsgBox "Could not change freeze panes." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Freeze Header Row"
End Sub

'------------------------------------------------------------
' SUBTOTAL sum (Ctrl+Shift+G)
'   the cell just below the cursor is taken as the header and skipped,
'   and everything from the next row to the end of the data goes into SUBTOTAL(9,).
'   select several cells across and each column gets its own.
'------------------------------------------------------------
Public Sub SubtotalSum()
    SubtotalCore 9, "SUBTOTAL Sum", False, "Ctrl+Shift+G"
End Sub

Public Sub SubtotalCount()
    SubtotalCore 3, "SUBTOTAL Count", True, "Ctrl+Alt+G"
End Sub

'------------------------------------------------------------
' the body shared by sum (9) and count (3)
'   funcNo : SUBTOTAL's first argument.  9 = SUM, 3 = COUNTA
'   forceFmt : True switches to #,##0 (for a count in a currency-formatted column)
'------------------------------------------------------------
Private Sub SubtotalCore(ByVal funcNo As Long, _
                     ByVal title2 As String, _
                     ByVal forceFormat As Boolean, _
                     ByVal keyLabel As String)
    Dim ws As Worksheet
    Dim sel As Range
    Dim c As Range
    Dim visible As Range, area As Range
    Dim startRow As Long, endRow As Long
    Dim addedCount As Long
    Dim oldScreen As Boolean
    Dim lastRow As Long
    Dim i As Long
    Dim totalRows As Long, visibleRows As Long
    Dim extraText As String
    Dim prevFormula() As Variant
    Dim prevBold() As Variant
    Dim prevFormat() As Variant

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, title2
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cells first.", vbExclamation, title2
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "Separate ranges cannot be processed at once.", vbExclamation, title2
        Exit Sub
    End If
    If sel.Rows.Count > 1 Then
        MsgBox "Please select a single row." & vbCrLf & _
               "Select the cells across and one will be put in each column.", _
               vbExclamation, title2
        Exit Sub
    End If

    ' starting one row past the header
    startRow = sel.Row + 2
    If startRow > ws.Rows.Count Then
        MsgBox "Could not find any data.", vbExclamation, title2
        Exit Sub
    End If

    ' last data row : the last row of the table block below the header
    '   a criteria block separated by a blank row is left out on its own
    endRow = LastRowOfBlock(ws, sel.Cells(1, 1).Offset(1, 0))

    If endRow < startRow Then
        endRow = LastRowOfBlock(ws, sel.Cells(1, 1).Offset(2, 0))
    End If

    ' fall back to the whole column when no table block is found (a plain table with no criteria block)
    If endRow < startRow Then
        For Each c In sel.Cells
            lastRow = LastRowOfColumn(ws, c.Column, startRow)
            If lastRow > endRow Then endRow = lastRow
        Next c
    End If

    If endRow < startRow Then
        MsgBox "There is no data to calculate." & vbCrLf & vbCrLf & _
               "Put the cursor where it should go and run again." & vbCrLf & _
               "The cell right below is treated as the header and skipped.", vbExclamation, title2
        Exit Sub
    End If

    '--- when it is the same spot as the last press of this key, put it back ---
    If mUndoOn And Len(mUndoAddr) > 0 And mUndoFunc = funcNo Then
        If StrComp(mUndoWb, ws.Parent.Name, vbTextCompare) = 0 And _
           StrComp(mUndoWs, ws.Name, vbTextCompare) = 0 And _
           StrComp(mUndoAddr, sel.Address, vbTextCompare) = 0 Then
            UndoSubtotal
            Exit Sub
        End If
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' keep the existing values so it can be undone
    ReDim prevFormula(1 To sel.Cells.Count)
    ReDim prevBold(1 To sel.Cells.Count)
    ReDim prevFormat(1 To sel.Cells.Count)
    i = 0
    For Each c In sel.Cells
        i = i + 1
        prevFormula(i) = c.Formula
        prevBold(i) = c.Font.Bold
        prevFormat(i) = c.NumberFormat
    Next c
    mUndoWb = ws.Parent.Name
    mUndoWs = ws.Name
    mUndoAddr = sel.Address
    mUndoFormula = prevFormula
    mUndoBold = prevBold
    mUndoNF = prevFormat
    mUndoFunc = funcNo
    mUndoTitle = title2
    mUndoOn = True

    For Each c In sel.Cells
        c.Formula = "=SUBTOTAL(" & funcNo & "," & _
                    ws.Cells(startRow, c.Column).Address(False, False) & ":" & _
                    ws.Cells(endRow, c.Column).Address(False, False) & ")"
        c.Font.Bold = True
        If forceFormat Then c.NumberFormat = "#,##0"
        addedCount = addedCount + 1
    Next c

    '--- for a count, report the total and visible row numbers too ---
    extraText = ""
    If funcNo = 3 Then
        totalRows = endRow - startRow + 1
        visibleRows = totalRows
        On Error Resume Next
        If ws.FilterMode Then
            Set visible = ws.Range(ws.Cells(startRow, sel.Cells(1, 1).Column), _
                                ws.Cells(endRow, sel.Cells(1, 1).Column)) _
                         .SpecialCells(xlCellTypeVisible)
            If Not visible Is Nothing Then
                visibleRows = 0
                For Each area In visible.Areas
                    visibleRows = visibleRows + area.Rows.Count
                Next area
            End If
        End If
        Err.Clear
        On Error GoTo Fail
        extraText = "  |  total " & Format$(totalRows, "#,##0") & " rows / visible " & _
                 Format$(visibleRows, "#,##0") & " rows"
    End If

    Application.ScreenUpdating = oldScreen

    ' register it so Ctrl+Z undoes it
    On Error Resume Next
    Application.OnUndo title2 & " - undo", MacroPrefix() & "UndoSubtotal"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = title2 & " " & addedCount & " cells added. (" & _
                            startRow & " to row " & endRow & " rows)" & extraText & "  |  " & _
                            "Ctrl+Z or " & keyLabel & " to undo"
    ScheduleStatusClear
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not insert it." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, title2
End Sub

'--- the last row of the whole table (CurrentRegion) ---
Private Function LastRowOfBlock(ByVal ws As Worksheet, ByVal anchorCell As Range) As Long
    Dim rg As Range
    On Error GoTo Fail
    Set rg = anchorCell.CurrentRegion
    If rg Is Nothing Then Exit Function
    LastRowOfBlock = rg.Row + rg.Rows.Count - 1
    Exit Function
Fail:
    LastRowOfBlock = 0
End Function

'--- the last data row in that column (walking up from the bottom) ---
Private Function LastRowOfColumn(ByVal ws As Worksheet, _
                               ByVal colNo As Long, _
                               ByVal startRow As Long) As Long
    Dim r As Long
    On Error GoTo Fail
    r = ws.Cells(ws.Rows.Count, colNo).End(xlUp).Row
    If r < startRow Then r = 0
    LastRowOfColumn = r
    Exit Function
Fail:
    LastRowOfColumn = 0
End Function

'------------------------------------------------------------
' undo - shared by SUBTOTAL sum and count
'------------------------------------------------------------
Public Sub UndoSubtotal()
    Dim ws As Worksheet
    Dim rg As Range
    Dim i As Long
    Dim oldScreen As Boolean
    Dim title2 As String
    Dim errNo As Long, errDesc As String

    title2 = mUndoTitle
    If Len(title2) = 0 Then title2 = "SUBTOTAL Sum"

    If Len(mUndoAddr) = 0 Then
        MsgBox "" & title2 & " not found.", vbInformation, title2
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set ws = Application.Workbooks(mUndoWb).Worksheets(mUndoWs)
    Set rg = ws.Range(mUndoAddr)

    For i = 1 To rg.Cells.Count
        If Len(CStr(mUndoFormula(i))) = 0 Then
            rg.Cells(i).ClearContents
        Else
            rg.Cells(i).Formula = mUndoFormula(i)
        End If
        rg.Cells(i).Font.Bold = mUndoBold(i)
        rg.Cells(i).NumberFormat = mUndoNF(i)
    Next i

    mUndoAddr = ""
    mUndoOn = False
    Application.ScreenUpdating = oldScreen

    Application.StatusBar = title2 & " undone."
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not undo." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, title2
End Sub

'------------------------------------------------------------
' undo the duplicate highlighting (called by Ctrl+Z)
'------------------------------------------------------------
Public Sub UndoDuplicateColors()
    Dim ws As Worksheet
    Dim rng As Range
    Dim r As Long, c As Long
    Dim colorSet As Object
    Dim blankAddr As String
    Dim keyColor As String
    Dim key2 As Variant
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If Len(mDupAddr) = 0 Or mDupCount = 0 Then
        MsgBox "There is no duplicate highlighting to undo.", vbInformation, "Highlight Duplicate Rows"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set ws = Application.Workbooks(mDupWb).Worksheets(mDupWs)
    Set rng = ws.Range(mDupAddr)
    Set colorSet = CreateObject("Scripting.Dictionary")

    ' gather addresses by color and restore them in one go
    For r = 1 To mDupCount
        If IsArray(mDupColor(r)) Then
            ' only rows that held mixed colors go cell by cell
            For c = 1 To UBound(mDupColor(r))
                With rng.Cells(mDupRows(r), c).Interior
                    If mDupIndex(r)(c) = xlColorIndexNone Then
                        .ColorIndex = xlColorIndexNone
                    Else
                        .Color = mDupColor(r)(c)
                    End If
                End With
            Next c
        ElseIf mDupIndex(r) = xlColorIndexNone Then
            If Len(blankAddr) > 0 Then blankAddr = blankAddr & ","
            blankAddr = blankAddr & rng.Rows(mDupRows(r)).Address(False, False)
            If Len(blankAddr) > 200 Then
                rng.Worksheet.Range(blankAddr).Interior.ColorIndex = xlColorIndexNone
                blankAddr = ""
            End If
        Else
            keyColor = CStr(mDupColor(r))
            If colorSet.Exists(keyColor) Then
                colorSet(keyColor) = colorSet(keyColor) & "," & rng.Rows(mDupRows(r)).Address(False, False)
            Else
                colorSet.Add keyColor, rng.Rows(mDupRows(r)).Address(False, False)
            End If
            If Len(colorSet(keyColor)) > 200 Then
                rng.Worksheet.Range(colorSet(keyColor)).Interior.Color = CLng(keyColor)
                colorSet(keyColor) = ""
            End If
        End If
    Next r

    If Len(blankAddr) > 0 Then rng.Worksheet.Range(blankAddr).Interior.ColorIndex = xlColorIndexNone
    For Each key2 In colorSet.Keys
        If Len(colorSet(key2)) > 0 Then
            rng.Worksheet.Range(colorSet(key2)).Interior.Color = CLng(key2)
        End If
    Next key2

    mDupAddr = ""
    mDupCount = 0
    Application.ScreenUpdating = oldScreen

    Application.StatusBar = "The duplicate highlighting was undone."
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not undo." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Highlight Duplicate Rows"
End Sub

'------------------------------------------------------------
' filter by the selected cells (Ctrl+Shift+Z)
'   turns the filter on over the current region when there is none, and uses the cursor's value as the criteria.
'------------------------------------------------------------
Public Sub FilterBySelectedValue()
    Dim ws As Worksheet
    Dim lo As Object
    Dim fr As Range
    Dim baseRange As Range
    Dim c As Range
    Dim sel As Range
    Dim v As Variant
    Dim crit As String
    Dim fld As Long
    Dim leftRows As Long
    Dim isOn As Boolean
    Dim colNo() As Long, valLists() As Variant
    Dim blankCount() As Long, valCount() As Long
    Dim colTotal As Long, i As Long, k As Long
    Dim alreadyOn As Long
    Dim summaryTxt As String
    Dim maxCount As Long
    Dim d As Object
    Dim keyArr As Variant
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Filter by Selected Value"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the cells to filter by, then run again.", vbExclamation, "Filter by Selected Value"
        Exit Sub
    End If
    Set sel = Selection
    If sel.Cells.CountLarge > 200000 Then
        MsgBox "The selection is too large." & vbCrLf & _
               "Please select only the cells holding the values you want to filter by.", vbExclamation, "Filter by Selected Value"
        Exit Sub
    End If
    Set c = ActiveCell

    On Error GoTo Fail

    '--- inside a table (ListObject), work against the table ---
    On Error Resume Next
    Set lo = c.ListObject
    Err.Clear
    On Error GoTo Fail

    If Not lo Is Nothing Then
        If Not lo.ShowAutoFilter Then lo.ShowAutoFilter = True
        Set fr = lo.Range
    Else
        Set baseRange = SkipNonHeaderRows(c.CurrentRegion)
        If baseRange Is Nothing Then
            MsgBox "Could not find any data.", vbExclamation, "Filter by Selected Value"
            Exit Sub
        End If
        If baseRange.Cells.CountLarge < 2 Or baseRange.Rows.Count < 2 Then
            MsgBox "Put the cursor on a cell with data and run again.", _
                   vbExclamation, "Filter by Selected Value"
            Exit Sub
        End If

        If ws.AutoFilterMode Then
            ' when the existing filter's header row is out of place, start a new one
            If ws.AutoFilter.Range.Row <> baseRange.Row Or _
               ws.AutoFilter.Range.Column <> baseRange.Column Then
                On Error Resume Next
                ws.ShowAllData
                Err.Clear
                On Error GoTo Fail
                ws.AutoFilterMode = False
                baseRange.AutoFilter
            End If
        Else
            baseRange.AutoFilter
        End If

        Set fr = ws.AutoFilter.Range
    End If

    '--- group the selected cells by column ---
    '    several in the same column = OR,  across columns = AND
    '    (that is how AutoFilter works anyway. each column is set on its own, and two of them are ANDed)
    ReDim colNo(1 To fr.Columns.Count)
    ReDim valLists(1 To fr.Columns.Count)
    ReDim blankCount(1 To fr.Columns.Count)
    ReDim valCount(1 To fr.Columns.Count)

    For Each c In sel.Cells
        fld = c.Column - fr.Column + 1
        If fld >= 1 And fld <= fr.Columns.Count And c.Row > fr.Row Then
            If c.Row <= fr.Row + fr.Rows.Count - 1 Then
                ' make a slot the first time this column comes up
                k = 0
                For i = 1 To colTotal
                    If colNo(i) = fld Then k = i: Exit For
                Next i
                If k = 0 Then
                    colTotal = colTotal + 1
                    k = colTotal
                    colNo(k) = fld
                    Set d = CreateObject("Scripting.Dictionary")
                    d.CompareMode = 1
                    Set valLists(k) = d
                End If

                v = c.Value
                If IsEmpty(v) Or Len(CStr(v)) = 0 Then
                    blankCount(k) = blankCount(k) + 1
                Else
                    If Not valLists(k).Exists(c.Text) Then
                        valLists(k).Add c.Text, 1
                        valCount(k) = valCount(k) + 1
                    End If
                End If
            End If
        End If
    Next c

    If colTotal = 0 Then
        MsgBox "Please select data cells inside the filter range." & vbCrLf & vbCrLf & _
               "The header row and anything outside the range are ignored." & vbCrLf & _
               "Filter range : " & fr.Address(False, False), _
               vbExclamation, "Filter by Selected Value"
        Exit Sub
    End If

    '--- when every picked column is already filtered, clear those columns (toggle) ---
    alreadyOn = 0
    On Error Resume Next
    For i = 1 To colTotal
        isOn = False
        isOn = ws.AutoFilter.Filters(colNo(i)).On
        If isOn Then alreadyOn = alreadyOn + 1
    Next i
    Err.Clear
    On Error GoTo Fail

    If alreadyOn = colTotal Then
        On Error Resume Next
        For i = 1 To colTotal
            fr.AutoFilter Field:=colNo(i)
            If Len(summaryTxt) > 0 Then summaryTxt = summaryTxt & " · "
            summaryTxt = summaryTxt & fr.Cells(1, colNo(i)).Value
        Next i
        Err.Clear
        On Error GoTo Fail
        Application.StatusBar = summaryTxt & " filter cleared.  |  clear all with Ctrl+Shift+A"
        ScheduleStatusClear
        Exit Sub
    End If

    '--- check how many ---
    maxCount = 0
    For i = 1 To colTotal
        If valCount(i) > maxCount Then maxCount = valCount(i)
    Next i

    If maxCount > 1000 Then
        MsgBox "One column has " & Format$(maxCount, "#,##0") & "." & vbCrLf & _
               "That is too many for AutoFilter." & vbCrLf & vbCrLf & _
               "Please use the Advanced Filter." & vbCrLf & _
               "Press Ctrl+Shift+R to build the criteria columns, then Ctrl+Shift+F.", _
               vbExclamation, "Filter by Selected Value"
        Exit Sub
    End If

    If maxCount > 100 Then
        summaryTxt = ""
        For i = 1 To colTotal
            If Len(summaryTxt) > 0 Then summaryTxt = summaryTxt & vbCrLf
            summaryTxt = summaryTxt & "   " & fr.Cells(1, colNo(i)).Value & " : " & _
                   Format$(valCount(i), "#,##0") & ""
        Next i
        If MsgBox("This is what will be filtered." & vbCrLf & vbCrLf & summaryTxt & vbCrLf & vbCrLf & _
                  "Do you want to continue?", vbOKCancel + vbQuestion, _
                  "Filter by Selected Value") <> vbOK Then Exit Sub
    End If

    '--- set it per column ---
    summaryTxt = ""
    For i = 1 To colTotal
        fld = colNo(i)
        If valCount(i) = 0 Then
            ' everything picked in that column is blank
            fr.AutoFilter Field:=fld, Criteria1:="="
            crit = "(blank)"
        Else
            ' filter by the list of visible values.
            '   give Criteria1 a string and Excel may read it as a date.
            '   xlFilterValues does not go through that reading.
            keyArr = valLists(i).Keys
            fr.AutoFilter Field:=fld, _
                          Criteria1:=keyArr, _
                          Operator:=xlFilterValues
            If valCount(i) = 1 Then
                crit = CStr(keyArr(0))
            Else
                crit = valCount(i) & ""
            End If
        End If
        If Len(summaryTxt) > 0 Then summaryTxt = summaryTxt & " · "
        summaryTxt = summaryTxt & fr.Cells(1, fld).Value & " = " & crit
        If valCount(i) > 0 And blankCount(i) > 0 Then
            summaryTxt = summaryTxt & "(blank " & blankCount(i) & " skipped)"
        End If
    Next i

    leftRows = VisibleRowCount(fr)

    If leftRows = 0 Then
        On Error Resume Next
        For i = 1 To colTotal
            fr.AutoFilter Field:=colNo(i)
        Next i
        Err.Clear
        On Error GoTo Fail
        MsgBox "No rows matched, so the filter was cleared." & vbCrLf & vbCrLf & _
               summaryTxt & vbCrLf & _
               "Filter range : " & fr.Address(False, False), _
               vbExclamation, "Filter by Selected Value"
        Exit Sub
    End If

    Application.StatusBar = summaryTxt & "  ->  " & Format$(leftRows, "#,##0") & "" & _
                            "  |  clear with Ctrl+Shift+A"
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "Could not apply the filter." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Filter by Selected Value"
End Sub

'--- so that * ? ~ are taken literally in the filter criteria ---
Private Function EscapeWildcards(ByVal s As String) As String
    s = Replace(s, "~", "~~")
    s = Replace(s, "*", "~*")
    s = Replace(s, "?", "~?")
    EscapeWildcards = s
End Function

'------------------------------------------------------------
' paste / remove the shortcut list (Ctrl+Shift+K)
'   sticks it to the right of the screen like a note. press again to remove it.
'------------------------------------------------------------
Public Sub CheatSheet()
    Dim ws As Worksheet
    Dim shp As Object
    Dim leftPart As Single, topPos As Single
    Dim vr As Range
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Shortcut List"
        Exit Sub
    End If
    Set ws = ActiveSheet

    On Error GoTo Fail

    ' remove it when it is already there
    On Error Resume Next
    Set shp = ws.Shapes(CHEAT_SHEET_SHAPE)
    Err.Clear
    On Error GoTo Fail

    If Not shp Is Nothing Then
        shp.Delete
        Application.StatusBar = "Removed the shortcut list."
        ScheduleStatusClear
        Exit Sub
    End If

    ' stick it to the top right of the screen
    leftPart = 40
    topPos = 40
    On Error Resume Next
    Set vr = ActiveWindow.VisibleRange
    If Not vr Is Nothing Then
        leftPart = vr.Left + vr.Width - 265
        topPos = vr.Top + 12
        If leftPart < vr.Left Then leftPart = vr.Left + 8
    End If
    Err.Clear
    On Error GoTo Fail

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPart, topPos, 300, 200)

    With shp
        .Name = CHEAT_SHEET_SHAPE
        .Fill.ForeColor.RGB = RGB(255, 253, 240)
        .Fill.Transparency = 0.05
        .Line.ForeColor.RGB = RGB(190, 190, 175)
        .Line.Weight = 1
        With .TextFrame
            .MarginLeft = 8
            .MarginRight = 8
            .MarginTop = 8
            .MarginBottom = 8
            .AutoSize = False
            With .Characters
                .Text = CheatSheetText()
                .Font.Name = "Segoe UI"
                .Font.Size = 9
                .Font.Color = RGB(35, 35, 40)
            End With
        End With
    End With

    ' make the title line stand out
    On Error Resume Next
    With shp.TextFrame.Characters(1, 14).Font
        .Bold = True
        .Size = 11
        .Color = RGB(0, 70, 150)
    End With
    Err.Clear
    On Error GoTo Fail

    HighlightShortcuts shp

    ' size the box to the text (so nothing is cut off)
    On Error Resume Next
    With shp.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
    End With
    Err.Clear
    On Error GoTo Fail

    ' nudge it so it does not fall off the screen
    On Error Resume Next
    If Not vr Is Nothing Then
        shp.Left = vr.Left + vr.Width - shp.Width - 12
        If shp.Left < vr.Left + 6 Then shp.Left = vr.Left + 6
        shp.Top = vr.Top + 6
    End If
    Err.Clear
    On Error GoTo Fail

    ' so Ctrl+Z removes it too
    mCheatWb = ws.Parent.Name
    mCheatWs = ws.Name
    On Error Resume Next
    Application.OnUndo "Remove the shortcut list", MacroPrefix() & "RemoveCheatSheet"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "Shortcut list pasted. Press Ctrl+Shift+K again to remove it."
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "Could not paste the shortcut list." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Shortcut List"
End Sub

Private Function CheatSheetText() As String
    Dim s As String
    s = "Excel Work Tools  v" & TOOL_VERSION & vbLf & vbLf
    s = s & "Ctrl+Shift+M    Work Tools menu  -  when in doubt, press this one" & vbLf & vbLf
    s = s & "[ Edit ]" & vbLf
    s = s & "Ctrl+Shift+O    Open New Window  (same file in two windows)" & vbLf
    s = s & "Ctrl+Shift+Q    Copy to New Workbook  (current sheet into a new workbook)" & vbLf
    s = s & "Ctrl+Shift+X    Comma Format  (set the negative and zero look in Ctrl+Shift+M)" & vbLf
    s = s & "Ctrl+Shift+Y    Text to Columns - as General  (text numbers become real numbers)" & vbLf
    s = s & "Ctrl+Shift+I    Text to Columns - as Text" & vbLf
    s = s & "Ctrl+Shift+U    Table Borders  (bold header + vertical dividers)" & vbLf
    s = s & "Ctrl+Alt+Z      Center Align / Default  (toggle)" & vbLf
    s = s & "Ctrl+Alt+X      Format Brush  (copy and paste with one key)" & vbLf
    s = s & "Ctrl+Alt+S      Center Across Selection  (centered without merging)" & vbLf
    s = s & "Ctrl+Shift+W    AutoFit rows and columns" & vbLf
    s = s & "Ctrl+Shift+S    Freeze / unfreeze header row  (header detected for you)" & vbLf
    s = s & "Ctrl+Shift+H    Tidy All Sheets  (every sheet to A1, zoom 100%)" & vbLf & vbLf
    s = s & "[ Fill & Select ]" & vbLf
    s = s & "Ctrl+Alt+A      Flash Fill  (like Ctrl+E, but to the end of the table with no range)" & vbLf
    s = s & "Ctrl+Alt+Q      Auto Fill Series  (1,2,3 / Mon,Tue,Wed to the end of the table)" & vbLf
    s = s & "Ctrl+Shift+1    Fill Formula Down  (to the end of the table)" & vbLf
    s = s & "Ctrl+Shift+4    Fill Blanks From Above  (tidy up unmerged tables)" & vbLf
    s = s & "Ctrl+Shift+2    Insert Index Column  (No, =ROW()-header row)" & vbLf
    s = s & "Ctrl+Shift+3    Select to End of Table  (ignores blanks)" & vbLf
    s = s & "Ctrl+Alt+E      Remove Rows Above Header  (tidy up before exporting)" & vbLf
    s = s & vbLf
    s = s & "[ Data ]" & vbLf
    s = s & "Ctrl+Shift+J    Create PivotTable  (new sheet, range detected)" & vbLf
    s = s & "Ctrl+Shift+T    Classic PivotTable cleanup  (removes subtotals, cannot be undone)" & vbLf
    s = s & "Ctrl+Shift+D    Highlight Duplicate Rows  (Ctrl-click cells to compare values instead)" & vbLf
    s = s & "Ctrl+Shift+V    Find Duplicates by Column  (pick columns, then highlight · number · extract)" & vbLf
    s = s & "Ctrl+Shift+L    Multi VLOOKUP  (group matching values and spread them out)" & vbLf
    s = s & "Ctrl+Alt+D      Compare Two Columns  (between lists · within each row)" & vbLf
    s = s & "Ctrl+Alt+B      Column Headers to Row Data  (Jan Feb Mar columns -> one column)" & vbLf
    s = s & "Ctrl+Shift+C    Clear Cell Fill  (removes fill color in the selection)" & vbLf
    s = s & "Ctrl+Shift+G    SUBTOTAL Sum  (in the cell above the header, press again to undo)" & vbLf
    s = s & "Ctrl+Alt+G      SUBTOTAL Count  (filled cells, follows the filter)" & vbLf & vbLf
    s = s & "[ Filter ]" & vbLf
    s = s & "Ctrl+Shift+A    Filter On / Off" & vbLf
    s = s & "Ctrl+Shift+Z    Filter by selected cells  (press again to clear those columns)" & vbLf
    s = s & "Ctrl+Shift+F    Run Advanced Filter  (press again to clear)" & vbLf
    s = s & "Ctrl+Shift+R    Set Up Criteria Columns  (headers copied for you)" & vbLf
    s = s & "        Advanced Filter needs one blank row below the data," & vbLf
    s = s & "        then a criteria block under it." & vbLf & vbLf
    s = s & "[ Export ]" & vbLf
    s = s & "Ctrl+Shift+E    Export Sheet as Text  (for txt -> dat conversion)" & vbLf
    s = s & "Ctrl+Shift+N    Copy File Path  (full path to the clipboard)" & vbLf & vbLf
    s = s & "[ Other ]" & vbLf
    s = s & "Ctrl+Alt+F      Find by Value  (matches results, selects every hit)" & vbLf
    s = s & "Ctrl+Shift+P    Set Password  (open password, default 1234)" & vbLf
    s = s & "Ctrl+Shift+K    Paste / remove this list" & vbLf & vbLf
    s = s & "When you close a saved file with no open password, you are asked about it." & vbLf & vbLf
    s = s & "Made by " & AUTHOR & " in Korea  ·  " & SITE
    CheatSheetText = s
End Function

'------------------------------------------------------------
' remove the shortcut list (called by Ctrl+Z)
'------------------------------------------------------------
Public Sub RemoveCheatSheet()
    Dim ws As Worksheet
    Dim shp As Object

    If Len(mCheatWs) = 0 Then Exit Sub

    On Error Resume Next
    Set ws = Application.Workbooks(mCheatWb).Worksheets(mCheatWs)
    If Not ws Is Nothing Then
        Set shp = ws.Shapes(CHEAT_SHEET_SHAPE)
        If Not shp Is Nothing Then shp.Delete
    End If
    Err.Clear
    On Error GoTo 0

    mCheatWs = ""
    Application.StatusBar = "Removed the shortcut list."
    ScheduleStatusClear
End Sub

'------------------------------------------------------------
' text to columns - as General (Ctrl+Shift+Y)
' text to columns - as Text (Ctrl+Shift+I)
'   every delimiter is off, so only the format is re-read. no column is split.
'   several columns can be done at once.
'------------------------------------------------------------
Public Sub TextToGeneral()
    SplitText xlGeneralFormat, "General"
End Sub

Public Sub TextToText()
    SplitText xlTextFormat, "Text"
End Sub

Private Sub SplitText(ByVal fmtMode As Long, ByVal fmtName As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim blk As Range
    Dim col As Range
    Dim noticeText As String
    Dim canUndo As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Text to Columns"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the cells to convert, then run again.", vbExclamation, "Text to Columns"
        Exit Sub
    End If

    '--- work out the range ---
    If Selection.Cells.CountLarge >= 2 Then
        Set rng = Selection
    Else
        Set blk = SkipNonHeaderRows(ActiveCell.CurrentRegion)
        If blk Is Nothing Then
            MsgBox "Could not find any data.", vbExclamation, "Text to Columns"
            Exit Sub
        End If
        Set rng = ws.Range(ws.Cells(blk.Row, ActiveCell.Column), _
                           ws.Cells(blk.Row + blk.Rows.Count - 1, ActiveCell.Column))
    End If

    If rng.Areas.Count > 1 Then
        MsgBox "Separate ranges cannot be processed at once." & vbCrLf & _
               "Select a single continuous range, then run again.", vbExclamation, "Text to Columns"
        Exit Sub
    End If

    ' guard against a whole column or row being selected
    On Error Resume Next
    Set rng = Application.Intersect(rng, ws.UsedRange)
    Err.Clear
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "There is no data to convert.", vbExclamation, "Text to Columns"
        Exit Sub
    End If

    canUndo = (rng.Cells.CountLarge <= 100000)

    '--- confirm before running ---
    noticeText = "The selection was read as [" & fmtName & "] format." & vbCrLf & vbCrLf & _
           "Range : " & rng.Address(False, False) & _
           "   (" & Format$(rng.Columns.Count, "#,##0") & " columns / " & _
           Format$(rng.Rows.Count, "#,##0") & " rows)" & vbCrLf & vbCrLf & _
           "· All delimiters are off, so columns will not be split." & vbCrLf

    If fmtMode = xlGeneralFormat Then
        noticeText = noticeText & "· Numbers stored as text become real numbers." & vbCrLf & _
                     "· Leading zeros are lost.  (00123 -> 123)" & vbCrLf
    Else
        noticeText = noticeText & "· Every value becomes text and can no longer be calculated." & vbCrLf
    End If

    noticeText = noticeText & IIf(canUndo, "· You can undo this with Ctrl+Z.", _
                                    "· The range is large, so Ctrl+Z is not supported here.") & _
           vbCrLf & vbCrLf & "Do you want to continue?"

    If MsgBox(noticeText, vbOKCancel + vbQuestion, "Text to Columns - " & fmtName) = vbCancel Then Exit Sub

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '--- kept so it can be undone ---
    If canUndo Then
        mTxtWb = ws.Parent.Name
        mTxtWs = ws.Name
        mTxtAddr = rng.Address
        mTxtValue = rng.Formula
        mTxtFormat = rng.NumberFormat
        mTxtOn = True
    Else
        mTxtOn = False
    End If

    '--- re-apply the format, column by column ---
    '   when going back to [General], a cell format left as [Text](@) means
    '   the value turns numeric while the ribbon still reads [Text].
    '   so the cell format is cleared to General first.
    If fmtMode = xlGeneralFormat Then
        On Error Resume Next
        rng.NumberFormat = "General"
        Err.Clear
        On Error GoTo Fail
    End If

    For Each col In rng.Columns
        col.TextToColumns Destination:=col.Cells(1, 1), _
                          DataType:=xlDelimited, _
                          TextQualifier:=xlDoubleQuote, _
                          ConsecutiveDelimiter:=False, _
                          Tab:=False, Semicolon:=False, Comma:=False, _
                          Space:=False, Other:=False, _
                          FieldInfo:=Array(1, fmtMode), _
                          TrailingMinusNumbers:=True
    Next col

    ' Excel remembers the last [Text to Columns] settings and
    ' applies them to later pastes as well.
    ' leave the delimiters off and pasting tab-separated values
    ' puts everything in one cell, so they are reset to the default (tab).
    RestoreTextToColumns

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    If canUndo Then
        On Error Resume Next
        Application.OnUndo "Undo Text to Columns", MacroPrefix() & "UndoSplitText"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "[" & fmtName & "] format.  " & _
                            rng.Address(False, False) & _
                            IIf(canUndo, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    RestoreTextToColumns
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not convert." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Text to Columns"
End Sub

'------------------------------------------------------------
' undo text to columns (called by Ctrl+Z)
'------------------------------------------------------------
Public Sub UndoSplitText()
    Dim ws As Worksheet
    Dim rng As Range
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim errNo As Long, errDesc As String

    If Not mTxtOn Or Len(mTxtAddr) = 0 Then
        MsgBox "There is no Text to Columns to undo.", vbInformation, "Text to Columns"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = Application.Workbooks(mTxtWb).Worksheets(mTxtWs)
    Set rng = ws.Range(mTxtAddr)

    rng.NumberFormat = mTxtFormat
    rng.Formula = mTxtValue

    mTxtOn = False
    mTxtAddr = ""

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    Application.StatusBar = "Text to Columns was undone."
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not undo." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Text to Columns"
End Sub

'--- in the shortcut list, bold only the Ctrl+Shift+? part ---
Private Sub HighlightShortcuts(ByVal shp As Object)
    Dim s As String
    Dim p As Long
    Dim len2 As Long

    On Error Resume Next
    s = shp.TextFrame.Characters.Text
    len2 = Len("Ctrl+Shift+X")

    p = InStr(1, s, "Ctrl+Shift+", vbTextCompare)
    Do While p > 0
        With shp.TextFrame.Characters(p, len2).Font
            .Bold = True
            .Color = RGB(0, 70, 150)
        End With
        p = InStr(p + 1, s, "Ctrl+Shift+", vbTextCompare)
    Loop

    ' make the whole menu line at the top stand out
    p = InStr(1, s, "Ctrl+Shift+M", vbTextCompare)
    If p > 0 Then
        shp.TextFrame.Characters(p, 50).Font.Bold = True
    End If

    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------
' undo - comma format
'------------------------------------------------------------
Public Sub UndoCommaFormat()
    On Error Resume Next
    If Not mFmtOn Or Len(mFmtAddr) = 0 Then
        MsgBox "There is no comma format to undo.", vbInformation, "Comma Format"
        Exit Sub
    End If
    Application.Workbooks(mFmtWb).Worksheets(mFmtWs).Range(mFmtAddr).NumberFormat = mFmtValue
    mFmtOn = False
    mFmtAddr = ""
    Err.Clear
    On Error GoTo 0
    Application.StatusBar = "The comma format was undone."
    ScheduleStatusClear
End Sub

'------------------------------------------------------------
' undo - AutoFit rows / columns
'------------------------------------------------------------
Public Sub UndoAutoFit()
    Dim rng As Range
    Dim i As Long
    Dim oldScreen As Boolean

    If Not mFitOn Or Len(mFitAddr) = 0 Then
        MsgBox "There is no AutoFit to undo.", vbInformation, "AutoFit Rows and Columns"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False
    Set rng = Application.Workbooks(mFitWb).Worksheets(mFitWs).Range(mFitAddr)
    For i = 1 To rng.Columns.Count
        rng.Columns(i).ColumnWidth = mFitColW(i)
    Next i
    For i = 1 To rng.Rows.Count
        rng.Rows(i).RowHeight = mFitRowH(i)
    Next i
    mFitOn = False
    mFitAddr = ""
    Err.Clear
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    Application.StatusBar = "Row and column sizes were restored."
    ScheduleStatusClear
End Sub

'------------------------------------------------------------
' undo - advanced filter criteria columns
'------------------------------------------------------------
Public Sub UndoCriteriaBlock()
    Dim rng As Range
    Dim r As Long, c As Long
    Dim oldScreen As Boolean

    If Not mCondOn Or Len(mCondAddr) = 0 Then
        MsgBox "There are no criteria columns to undo.", vbInformation, "Set Up Criteria Columns"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False

    Set rng = Application.Workbooks(mCondWb).Worksheets(mCondWs).Range(mCondAddr)

    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            With rng.Cells(r, c)
                .NumberFormat = mCondN(r, c)
                .Formula = mCondF(r, c)
                If mCondI(r, c) = xlColorIndexNone Then
                    .Interior.ColorIndex = xlColorIndexNone
                Else
                    .Interior.Color = mCondC(r, c)
                End If
                .Font.Bold = mCondB(r, c)
                .Font.Color = mCondFC(r, c)
            End With
        Next c
    Next r

    mCondOn = False
    mCondAddr = ""
    Err.Clear
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    Application.StatusBar = "The criteria columns were undone."
    ScheduleStatusClear
End Sub

'--- kept so the criteria columns can be undone : values, formats and colors together ---
Private Sub SaveCriteriaBlock(ByVal rg As Range)
    Dim r As Long, c As Long
    Dim nR As Long, nC As Long
    Dim vF() As Variant, vN() As Variant
    Dim vC() As Variant, vI() As Variant
    Dim vB() As Variant, vFC() As Variant

    mCondOn = False
    If rg Is Nothing Then Exit Sub
    If rg.Cells.CountLarge > 20000 Then Exit Sub

    nR = rg.Rows.Count
    nC = rg.Columns.Count

    On Error GoTo Fail
    ReDim vF(1 To nR, 1 To nC)
    ReDim vN(1 To nR, 1 To nC)
    ReDim vC(1 To nR, 1 To nC)
    ReDim vI(1 To nR, 1 To nC)
    ReDim vB(1 To nR, 1 To nC)
    ReDim vFC(1 To nR, 1 To nC)

    For r = 1 To nR
        For c = 1 To nC
            With rg.Cells(r, c)
                vF(r, c) = .Formula
                vN(r, c) = .NumberFormat
                vC(r, c) = .Interior.Color
                vI(r, c) = .Interior.ColorIndex
                vB(r, c) = .Font.Bold
                vFC(r, c) = .Font.Color
            End With
        Next c
    Next r

    mCondWb = rg.Worksheet.Parent.Name
    mCondWs = rg.Worksheet.Name
    mCondAddr = rg.Address
    mCondF = vF
    mCondN = vN
    mCondC = vC
    mCondI = vI
    mCondB = vB
    mCondFC = vFC
    mCondOn = True
    Exit Sub
Fail:
    mCondOn = False
End Sub

'------------------------------------------------------------
' open a new window (Ctrl+Shift+O)
'   shows the same file in two windows. (same as View - New Window)
'   to finish, just close the window.
'------------------------------------------------------------
Public Sub OpenNewWindow()
    Dim wb As Workbook
    Dim windowCount As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "Please open a file first.", vbExclamation, "Open New Window"
        Exit Sub
    End If

    Set wb = ActiveWorkbook

    On Error GoTo Fail
    ActiveWindow.NewWindow

    On Error Resume Next
    windowCount = wb.Windows.Count
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "New window opened.  (" & wb.Name & " window " & windowCount & ")" & _
                            "   Just close the window when you are done."
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "Could not open a new window." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Open New Window"
End Sub

'------------------------------------------------------------
' create a PivotTable (Ctrl+Shift+J)
'   builds an empty PivotTable on a new sheet from the current region.
'   drag the fields in from the field list yourself, then tidy up with Ctrl+Shift+T.
'------------------------------------------------------------
Public Sub CreatePivot()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim source As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim name2 As String
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Create PivotTable"
        Exit Sub
    End If
    Set ws = ActiveSheet

    '--- source range : the total row is left out ---
    Set source = SkipNonHeaderRows(ActiveCell.CurrentRegion)
    If source Is Nothing Then
        MsgBox "Could not find any data.", vbExclamation, "Create PivotTable"
        Exit Sub
    End If
    If source.Rows.Count < 2 Then
        MsgBox "A header row and data rows are both required.", vbExclamation, "Create PivotTable"
        Exit Sub
    End If

    On Error GoTo Fail

    ' lock the screen, recalculation and events (very slow on a big table otherwise)
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    '--- name the new sheet : Pivot_sourcesheetname ---
    name2 = PivotSheetName(ws.Parent, ws.Name)

    ' put it right in front of the source sheet
    Set newSheet = ws.Parent.Worksheets.Add(Before:=ws)
    On Error Resume Next
    newSheet.Name = name2
    Err.Clear
    On Error GoTo Fail

    Set pc = ws.Parent.PivotCaches.Create(SourceType:=xlDatabase, _
                                          SourceData:=source)
    Set pt = pc.CreatePivotTable(TableDestination:=newSheet.Range("B3"))

    ' note where the source is in A1
    On Error Resume Next
    With newSheet.Range("A1")
        .Value = "Source : " & ws.Name & "!" & source.Address(False, False)
        .Font.Size = 9
        .Font.Color = RGB(110, 110, 115)
    End With
    newSheet.Columns("A").ColumnWidth = 3
    Err.Clear
    On Error GoTo Fail

    newSheet.Activate
    newSheet.Range("B3").Select

    ' unlock everything, then show the field list
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    ActiveWorkbook.ShowPivotTableFieldList = True
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "PivotTable created.  Source " & ws.Name & "!" & _
                            source.Address(False, False) & _
                            "   Add the fields, then press Ctrl+Shift+T to tidy it up."
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
    MsgBox "Could not create the PivotTable." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Create PivotTable"
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal name2 As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(name2)
    Err.Clear
    On Error GoTo 0
    SheetExists = Not (sh Is Nothing)
End Function

'------------------------------------------------------------
' reset the settings
'   run before distributing to put everything back to factory defaults.
'   ask on close = on,  default password = 1234
'------------------------------------------------------------
Public Sub ResetSettings()
    Dim currentAuto As String
    Dim currentPwd As String
    Dim shp As Object
    Dim ws As Worksheet
    Dim removedCount As Long

    currentAuto = IIf(AskOnCloseEnabled(), "on", "off")
    currentPwd = ReadDefaultPassword()

    If MsgBox("This restores the settings to their defaults." & vbCrLf & vbCrLf & _
              "   Ask on close    : " & currentAuto & "  ->  on" & vbCrLf & _
              "   Default password  : " & currentPwd & "  ->  " & DEFAULT_PASSWORD & vbCrLf & _
              "   Comma format    : " & CommaFormatOneLine() & vbCrLf & _
              "                  ->  3,333 / △3,333 / -" & vbCrLf & vbCrLf & _
              "The shortcut list (Ctrl+Shift+K) in open files is removed too." & vbCrLf & vbCrLf & _
              "Do you want to continue?", vbOKCancel + vbQuestion, "Reset Settings") = vbCancel Then Exit Sub

    WriteSetting "auto", "1"
    WriteSetting "pwd", DEFAULT_PASSWORD
    WriteSetting "cmaColor", "R"
    WriteSetting "cmaNeg", "T"
    WriteSetting "cmaZero", "D"
    WriteSetting "cmaPlus", "N"
    WriteSetting "cmaPColor", "K"

    ' remove the shortcut-list shape from every open sheet
    On Error Resume Next
    For Each ws In Application.ActiveWorkbook.Worksheets
        Set shp = Nothing
        Set shp = ws.Shapes(CHEAT_SHEET_SHAPE)
        If Not shp Is Nothing Then
            shp.Delete
            removedCount = removedCount + 1
        End If
    Next ws
    Err.Clear
    On Error GoTo 0

    MsgBox "Settings restored to their defaults." & vbCrLf & vbCrLf & _
           "   Ask on close    : on" & vbCrLf & _
           "   Default password  : " & DEFAULT_PASSWORD & vbCrLf & _
           "   Comma format    : 3,333 / ▲3,333 / -" & vbCrLf & _
           IIf(removedCount > 0, "   Shortcut list " & removedCount & " removed" & vbCrLf, "") & vbCrLf & _
           "Close Excel completely before copying this for distribution." & vbCrLf & _
           "Copied to : " & ThisWorkbook.FullName, _
           vbInformation, "Reset Settings"
End Sub

'------------------------------------------------------------
' table borders (Ctrl+Shift+U)
'   header : bold + centered + boxed
'   data : vertical lines only (no horizontal lines between rows)
'   the outside border and the bottom line are drawn. all the same weight.
'------------------------------------------------------------
Public Sub TableBorders()
    Dim ws As Worksheet
    Dim rng As Range
    Dim header As Range
    Dim body As Range
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Table Borders"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the range to turn into a table, then run again.", vbExclamation, "Table Borders"
        Exit Sub
    End If

    If Selection.Cells.CountLarge >= 2 Then
        Set rng = Selection
    Else
        Set rng = SkipNonHeaderRows(ActiveCell.CurrentRegion)
    End If

    If rng Is Nothing Then
        MsgBox "Could not find data to turn into a table.", vbExclamation, "Table Borders"
        Exit Sub
    End If
    If rng.Areas.Count > 1 Then
        MsgBox "Separate ranges cannot be processed at once.", vbExclamation, "Table Borders"
        Exit Sub
    End If
    If rng.Rows.Count < 2 Then
        MsgBox "A header row and data rows are both required.", vbExclamation, "Table Borders"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    SaveBorders rng

    Set header = rng.Rows(1)
    Set body = rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count)

    '--- clear the existing borders ---
    rng.Borders.LineStyle = xlNone

    '--- the outside border ---
    DrawBorder rng, xlEdgeTop
    DrawBorder rng, xlEdgeBottom
    DrawBorder rng, xlEdgeLeft
    DrawBorder rng, xlEdgeRight

    '--- the vertical lines (all of them) ---
    If rng.Columns.Count > 1 Then DrawBorder rng, xlInsideVertical

    '--- the horizontal line below the header ---
    DrawBorder header, xlEdgeBottom

    '--- the header format ---
    header.Font.Bold = True
    header.HorizontalAlignment = xlCenter

    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    Application.OnUndo "Undo Table Borders", MacroPrefix() & "UndoTableBorders"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "Table borders applied. (" & rng.Address(False, False) & ")" & _
                            "  |  Ctrl+Z to undo"
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not apply the table borders." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Table Borders"
End Sub

Private Sub DrawBorder(ByVal rg As Range, ByVal pos As Long)
    On Error Resume Next
    With rg.Borders(pos)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Err.Clear
    On Error GoTo 0
End Sub

'--- kept so the table borders can be undone ---
Private Sub SaveBorders(ByVal rg As Range)
    Dim r As Long, c As Long, i As Long
    Dim nR As Long, nC As Long
    Dim vL() As Variant, vB() As Variant, vA() As Variant
    Dim pos As Variant

    mBdrOn = False
    If rg Is Nothing Then Exit Sub
    If rg.Cells.CountLarge > 20000 Then Exit Sub

    nR = rg.Rows.Count
    nC = rg.Columns.Count
    pos = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)

    On Error GoTo Fail
    ReDim vL(1 To nR, 1 To nC, 0 To 3)
    ReDim vB(1 To nR, 1 To nC)
    ReDim vA(1 To nR, 1 To nC)

    For r = 1 To nR
        For c = 1 To nC
            For i = 0 To 3
                vL(r, c, i) = rg.Cells(r, c).Borders(pos(i)).LineStyle
            Next i
            vB(r, c) = rg.Cells(r, c).Font.Bold
            vA(r, c) = rg.Cells(r, c).HorizontalAlignment
        Next c
    Next r

    mBdrWb = rg.Worksheet.Parent.Name
    mBdrWs = rg.Worksheet.Name
    mBdrAddr = rg.Address
    mBdrLine = vL
    mBdrBold = vB
    mBdrAlign = vA
    mBdrOn = True
    Exit Sub
Fail:
    mBdrOn = False
End Sub

'------------------------------------------------------------
' undo the table borders (called by Ctrl+Z)
'------------------------------------------------------------
Public Sub UndoTableBorders()
    Dim rng As Range
    Dim r As Long, c As Long, i As Long
    Dim pos As Variant
    Dim oldScreen As Boolean

    If Not mBdrOn Or Len(mBdrAddr) = 0 Then
        MsgBox "There are no table borders to undo.", vbInformation, "Table Borders"
        Exit Sub
    End If

    pos = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
    oldScreen = Application.ScreenUpdating

    On Error Resume Next
    Application.ScreenUpdating = False
    Set rng = Application.Workbooks(mBdrWb).Worksheets(mBdrWs).Range(mBdrAddr)

    rng.Borders.LineStyle = xlNone

    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            For i = 0 To 3
                If mBdrLine(r, c, i) <> xlNone Then
                    With rng.Cells(r, c).Borders(pos(i))
                        .LineStyle = mBdrLine(r, c, i)
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
                End If
            Next i
            rng.Cells(r, c).Font.Bold = mBdrBold(r, c)
            rng.Cells(r, c).HorizontalAlignment = mBdrAlign(r, c)
        Next c
    Next r

    mBdrOn = False
    mBdrAddr = ""
    Err.Clear
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    Application.StatusBar = "The table borders were undone."
    ScheduleStatusClear
End Sub

'--- kept so the cleared cell fill can be undone ---
Private Sub SaveFillColors(ByVal rg As Range)
    Dim ar As Range
    Dim cel As Range
    Dim i As Long, n As Long
    Dim vC() As Variant, vI() As Variant

    mBgOn = False
    If rg Is Nothing Then Exit Sub
    If rg.Cells.CountLarge > 50000 Then Exit Sub

    On Error GoTo Fail
    n = 0
    For Each ar In rg.Areas
        n = n + ar.Cells.Count
    Next ar
    If n = 0 Then Exit Sub

    ReDim vC(1 To n)
    ReDim vI(1 To n)

    i = 0
    For Each ar In rg.Areas
        For Each cel In ar.Cells
            i = i + 1
            vC(i) = cel.Interior.Color
            vI(i) = cel.Interior.ColorIndex
        Next cel
    Next ar

    mBgWb = rg.Worksheet.Parent.Name
    mBgWs = rg.Worksheet.Name
    mBgAddr = rg.Address
    mBgColor = vC
    mBgIndex = vI
    mBgOn = True
    Exit Sub
Fail:
    mBgOn = False
End Sub

'------------------------------------------------------------
' undo clearing the cell fill (called by Ctrl+Z)
'------------------------------------------------------------
Public Sub UndoFillColors()
    Dim rng As Range
    Dim ar As Range
    Dim cel As Range
    Dim i As Long
    Dim oldScreen As Boolean

    If Not mBgOn Or Len(mBgAddr) = 0 Then
        MsgBox "There is no fill color to undo.", vbInformation, "Clear Cell Fill"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False

    Set rng = Application.Workbooks(mBgWb).Worksheets(mBgWs).Range(mBgAddr)

    i = 0
    For Each ar In rng.Areas
        For Each cel In ar.Cells
            i = i + 1
            If mBgIndex(i) = xlColorIndexNone Then
                cel.Interior.ColorIndex = xlColorIndexNone
            Else
                cel.Interior.Color = mBgColor(i)
            End If
        Next cel
    Next ar

    mBgOn = False
    mBgAddr = ""
    Err.Clear
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    Application.StatusBar = "The cell fill color was restored."
    ScheduleStatusClear
End Sub

'--- build the pivot sheet name : Pivot_sourcename (31-char limit, no clashes) ---
Private Function PivotSheetName(ByVal wb As Workbook, ByVal sourceName As String) As String
    Dim baseName As String
    Dim candidate As String
    Dim suffix As String
    Dim i As Long

    ' strip the characters a sheet name cannot hold
    sourceName = Replace(sourceName, ":", "")
    sourceName = Replace(sourceName, "\", "")
    sourceName = Replace(sourceName, "/", "")
    sourceName = Replace(sourceName, "?", "")
    sourceName = Replace(sourceName, "*", "")
    sourceName = Replace(sourceName, "[", "")
    sourceName = Replace(sourceName, "]", "")

    baseName = "Pivot_" & sourceName
    If Len(baseName) > 31 Then baseName = Left$(baseName, 31)

    If Not SheetExists(wb, baseName) Then
        PivotSheetName = baseName
        Exit Function
    End If

    ' when one exists, add a number and trim to stay within 31 characters
    For i = 2 To 999
        suffix = CStr(i)
        candidate = baseName
        If Len(candidate) + Len(suffix) > 31 Then
            candidate = Left$(candidate, 31 - Len(suffix))
        End If
        candidate = candidate & suffix
        If Not SheetExists(wb, candidate) Then
            PivotSheetName = candidate
            Exit Function
        End If
    Next i

    PivotSheetName = "Pivot_" & Format$(Now, "hhnnss")
End Function

'------------------------------------------------------------
' find duplicate values (when several spots are picked for Ctrl+Shift+D)
'   compares the picked cells by value rather than by row, and gives matching ones the same color.
'------------------------------------------------------------
Private Sub FindDuplicateValues(ByVal rng As Range)
    Dim dic As Object
    Dim groupColor As Object
    Dim palette As Variant
    Dim ar As Range, cel As Range
    Dim key As String
    Dim groupCount As Long
    Dim checked As Long, paintedCount As Long
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If rng Is Nothing Then Exit Sub

    checked = 0
    On Error Resume Next
    checked = rng.Cells.CountLarge
    Err.Clear
    On Error GoTo 0

    If checked < 2 Then
        MsgBox "Please select two or more cells.", vbExclamation, "Find Duplicate Values"
        Exit Sub
    End If
    If checked > 50000 Then
        MsgBox "Too many cells selected. (" & Format$(checked, "#,##0") & " cells)" & vbCrLf & _
               "Please narrow it to 50,000 cells or fewer.", vbExclamation, "Find Duplicate Values"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set groupColor = CreateObject("Scripting.Dictionary")
    groupColor.CompareMode = 1

    palette = ColorPalette()

    ' pass 1 : count per value
    For Each ar In rng.Areas
        For Each cel In ar.Cells
            key = CStr(cel.Value)
            If Len(key) > 0 Then
                If dic.Exists(key) Then
                    dic(key) = dic(key) + 1
                Else
                    dic.Add key, 1
                End If
            End If
        Next cel
    Next ar

    ' kept so it can be undone
    SaveFillColors rng

    ' pass 2 : a color per group for values appearing twice or more
    For Each ar In rng.Areas
        For Each cel In ar.Cells
            key = CStr(cel.Value)
            If Len(key) > 0 Then
                If dic(key) >= 2 Then
                    If Not groupColor.Exists(key) Then
                        groupColor.Add key, palette(groupCount Mod (UBound(palette) + 1))
                        groupCount = groupCount + 1
                    End If
                    cel.Interior.Color = groupColor(key)
                    paintedCount = paintedCount + 1
                End If
            End If
        Next cel
    Next ar

    Application.ScreenUpdating = oldScreen

    If paintedCount > 0 And mBgOn Then
        On Error Resume Next
        Application.OnUndo "Undo Find Duplicate Values", MacroPrefix() & "UndoFillColors"
        Err.Clear
        On Error GoTo Fail
    End If

    If paintedCount = 0 Then
        MsgBox "selected " & Format$(checked, "#,##0") & " cells hold no matching values.", _
               vbInformation, "Find Duplicate Values"
    Else
        MsgBox "selected " & Format$(checked, "#,##0") & " cells were compared by value." & vbCrLf & _
               "Same values " & groupCount & " kinds / " & Format$(paintedCount, "#,##0") & " cells were marked." & _
               IIf(groupCount > UBound(palette) + 1, vbCrLf & "There are " & (UBound(palette) + 1) & _
                   " kinds, so some colors repeat.", "") & vbCrLf & vbCrLf & _
               "You can undo this with Ctrl+Z." & vbCrLf & _
               "To clear only the colors : select the range, then Ctrl+Shift+C", _
               vbInformation, "Find Duplicate Values"
    End If
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "No duplicate values were found." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Find Duplicate Values"
End Sub

'============================================================
'  find duplicates by key column (Ctrl+Shift+V)
'    pick the columns to compare, and rows whose values match across them are highlighted.
'    like [Remove Duplicates], but it only marks them instead of deleting.
'============================================================

Public Sub ColumnDuplicates()
    Dim ws As Worksheet
    Dim data As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Find Duplicates by Column"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set data = PickColumnRange(ws)
    If data Is Nothing Then Exit Sub

    ' reset the picked columns when the table has changed
    If StrComp(mColWs, ws.Name, vbTextCompare) <> 0 Or _
       StrComp(mColAddr, data.Address, vbTextCompare) <> 0 Then
        ReDim mColPick(1 To data.Columns.Count)
        mColWs = ws.Name
        mColAddr = data.Address
    End If

    On Error Resume Next
    If UBound(mColPick) <> data.Columns.Count Then
        ReDim mColPick(1 To data.Columns.Count)
    End If
    Err.Clear
    On Error GoTo 0

    ColumnPickMenu data
End Sub

'--- work out the table range ---
Private Function PickColumnRange(ByVal ws As Worksheet) As Range
    Dim rg As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell inside the table, then run again.", vbExclamation, "Find Duplicates by Column"
        Exit Function
    End If

    If Selection.Cells.CountLarge >= 2 And Selection.Areas.Count = 1 Then
        Set rg = Selection
    Else
        Set rg = SkipNonHeaderRows(ActiveCell.CurrentRegion)
    End If

    If rg Is Nothing Then
        MsgBox "Could not find a table.", vbExclamation, "Find Duplicates by Column"
        Exit Function
    End If
    If rg.Rows.Count < 3 Then
        MsgBox "One header row and at least two data rows are required.", vbExclamation, "Find Duplicates by Column"
        Exit Function
    End If
    If rg.Columns.Count > 60 Then
        MsgBox "Columns " & rg.Columns.Count & ", which is too many." & vbCrLf & _
               "Please select only the range holding the columns to compare.", vbExclamation, "Find Duplicates by Column"
        Exit Function
    End If

    Set PickColumnRange = rg
End Function

'--- the column-picking popup menu ---
Private Sub ColumnPickMenu(ByVal data As Range)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim i As Long
    Dim pickedCount As Long
    Dim name2 As String

    On Error Resume Next
    Application.CommandBars(BAR_COLUMN_PICK).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=BAR_COLUMN_PICK, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, _
                                         Temporary:=True)

    For i = 1 To data.Columns.Count
        If mColPick(i) Then pickedCount = pickedCount + 1
    Next i

    AddMenuTitle cb, "Find Duplicates by Column        (" & data.Address(False, False) & ")"
    AddMenuTitle cb, "■  Finds rows where every marked column has the same value."

    For i = 1 To data.Columns.Count
        name2 = Trim$(CStr(data.Cells(1, i).Value))
        If Len(name2) = 0 Then name2 = "(no header) " & Split(data.Cells(1, i).Address(True, False), "$")(0)

        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mColPick(i), "■  ", "□  ") & name2
        btn.OnAction = MacroPrefix() & "ColPickToggle" & Format$(i, "00")
        btn.Style = msoButtonCaption
        If i = 1 Then btn.BeginGroup = True
    Next i

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Highlight + group numbers   (both at once)"
    btn.OnAction = MacroPrefix() & "HighlightAndNumberDuplicates"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True
    btn.Enabled = (pickedCount > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Highlight duplicates   (" & pickedCount & " selected)"
    btn.OnAction = MacroPrefix() & "RunColumnDuplicates"
    btn.Style = msoButtonCaption
    btn.Enabled = (pickedCount > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Add duplicate group numbers   (2 columns at the right of the table)"
    btn.OnAction = MacroPrefix() & "AddDuplicateGroupNo"
    btn.Style = msoButtonCaption
    btn.Enabled = (pickedCount > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Duplicate rows only, to a new sheet   (all columns)"
    btn.OnAction = MacroPrefix() & "ExtractDuplicateRows"
    btn.Style = msoButtonCaption
    btn.Enabled = (pickedCount > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Unique list of picked columns to a new sheet   (duplicates removed)"
    btn.OnAction = MacroPrefix() & "ExtractUniqueList"
    btn.Style = msoButtonCaption
    btn.Enabled = (pickedCount > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Clear all"
    btn.OnAction = MacroPrefix() & "ColumnPickClearAll"
    btn.Style = msoButtonCaption

    cb.ShowPopup
    Exit Sub

Fail:
    MsgBox "Could not show the menu." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Find Duplicates by Column"
End Sub

Public Sub ColumnPickClearAll()
    Dim i As Long
    On Error Resume Next
    For i = LBound(mColPick) To UBound(mColPick)
        mColPick(i) = False
    Next i
    Err.Clear
    On Error GoTo 0
    ColumnDuplicates
End Sub

'--- flip one column on or off, then reopen the menu ---
Private Sub ColPickToggle(ByVal no As Long)
    On Error Resume Next
    mColPick(no) = Not mColPick(no)
    Err.Clear
    On Error GoTo 0
    ColumnDuplicates
End Sub

'--- the duplicate search itself ---
Public Sub RunColumnDuplicates()
    Dim ws As Worksheet
    Dim data As Range
    Dim dic As Object, groupColor As Object
    Dim palette As Variant
    Dim arr As Variant
    Dim visibleRange As Range, area As Range
    Dim visible() As Boolean
    Dim targetRows() As Variant
    Dim prevColor() As Variant, prevColorIndex() As Variant
    Dim cellColor() As Variant, cellColorIndex() As Variant
    Dim key As String
    Dim r As Long, c As Long, i As Long, startIdx As Long
    Dim rowCnt As Long, colCnt As Long
    Dim visibleCount As Long
    Dim colorAddr() As String
    Dim colorNo As Long
    Dim groupCount As Long, dupRows As Long
    Dim pickedNames As String
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = ActiveSheet
    Set data = ws.Range(mColAddr)
    rowCnt = data.Rows.Count
    colCnt = data.Columns.Count

    '--- work out the visible rows (header row left out) ---
    ReDim visible(1 To rowCnt)
    On Error Resume Next
    Set visibleRange = data.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo 0

    If visibleRange Is Nothing Then
        For r = 2 To rowCnt
            visible(r) = True
            visibleCount = visibleCount + 1
        Next r
    Else
        For Each area In visibleRange.Areas
            startIdx = area.Row - data.Row + 1
            For r = startIdx To startIdx + area.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not visible(r) Then
                        visible(r) = True
                        visibleCount = visibleCount + 1
                    End If
                End If
            Next r
        Next area
    End If

    If visibleCount < 2 Then
        MsgBox "Fewer than two visible data rows.", vbExclamation, "Find Duplicates by Column"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = data.Value

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set groupColor = CreateObject("Scripting.Dictionary")
    groupColor.CompareMode = 1
    palette = ColorPalette()

    ' pass 1 : count by the picked column combination
    For r = 2 To rowCnt
      If visible(r) Then
        key = ""
        For c = 1 To colCnt
            If mColPick(c) Then key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic.Exists(key) Then
            dic(key) = dic(key) + 1
        Else
            dic.Add key, 1
        End If
      End If
    Next r

    ' pass 2 : collect the rows involved
    ReDim targetRows(1 To rowCnt)
    For r = 2 To rowCnt
      If visible(r) Then
        key = ""
        For c = 1 To colCnt
            If mColPick(c) Then key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            dupRows = dupRows + 1
            targetRows(dupRows) = r
        End If
      End If
    Next r

    ' kept so it can be undone
    If dupRows > 0 Then
        ' read each row's color in one go. only mixed (Null) rows go cell by cell.
        ReDim prevColor(1 To dupRows)
        ReDim prevColorIndex(1 To dupRows)
        For r = 1 To dupRows
            With data.Rows(targetRows(r))
                prevColor(r) = .Interior.Color
                prevColorIndex(r) = .Interior.ColorIndex
            End With
            If IsNull(prevColor(r)) Or IsNull(prevColorIndex(r)) Then
                ReDim cellColor(1 To colCnt)
                ReDim cellColorIndex(1 To colCnt)
                For c = 1 To colCnt
                    cellColor(c) = data.Cells(targetRows(r), c).Interior.Color
                    cellColorIndex(c) = data.Cells(targetRows(r), c).Interior.ColorIndex
                Next c
                prevColor(r) = cellColor
                prevColorIndex(r) = cellColorIndex
            End If
        Next r
        mDupWb = ws.Parent.Name
        mDupWs = ws.Name
        mDupAddr = data.Address
        mDupRows = targetRows
        mDupColor = prevColor
        mDupIndex = prevColorIndex
        mDupCount = dupRows
    End If

    ' pass 3 : gather addresses per color and paint in one go
    dupRows = 0
    ReDim colorAddr(0 To UBound(palette))
    For r = 2 To rowCnt
      If visible(r) Then
        key = ""
        For c = 1 To colCnt
            If mColPick(c) Then key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            If Not groupColor.Exists(key) Then
                groupColor.Add key, groupCount Mod (UBound(palette) + 1)
                groupCount = groupCount + 1
            End If
            colorNo = groupColor(key)
            If Len(colorAddr(colorNo)) > 0 Then colorAddr(colorNo) = colorAddr(colorNo) & ","
            colorAddr(colorNo) = colorAddr(colorNo) & data.Rows(r).Address(False, False)
            dupRows = dupRows + 1
            If Len(colorAddr(colorNo)) > 200 Then
                ws.Range(colorAddr(colorNo)).Interior.Color = palette(colorNo)
                colorAddr(colorNo) = ""
            End If
        End If
      End If
    Next r

    For i = 0 To UBound(palette)
        If Len(colorAddr(i)) > 0 Then ws.Range(colorAddr(i)).Interior.Color = palette(i)
    Next i

    Application.ScreenUpdating = oldScreen

    If dupRows > 0 Then
        On Error Resume Next
        Application.OnUndo "Undo Find Duplicates by Column", MacroPrefix() & "UndoDuplicateColors"
        Err.Clear
        On Error GoTo Fail
    End If

    ' gather the names of the picked columns
    For c = 1 To colCnt
        If mColPick(c) Then
            If Len(pickedNames) > 0 Then pickedNames = pickedNames & " + "
            pickedNames = pickedNames & Trim$(CStr(data.Cells(1, c).Value))
        End If
    Next c

    If mQuiet Then Exit Sub

    If dupRows = 0 Then
        MsgBox "No duplicate rows found." & vbCrLf & vbCrLf & _
               "Key columns : " & pickedNames & vbCrLf & _
               "Checked range : " & data.Address(False, False) & _
               IIf(visibleCount < rowCnt - 1, "  (visible " & visibleCount & " rows only)", ""), _
               vbInformation, "Find Duplicates by Column"
    Else
        MsgBox "Duplicates " & groupCount & " kinds / " & Format$(dupRows, "#,##0") & " rows highlighted." & vbCrLf & vbCrLf & _
               "Key columns : " & pickedNames & vbCrLf & _
               "Checked range : " & data.Address(False, False) & _
               IIf(visibleCount < rowCnt - 1, "  (visible " & visibleCount & " rows only)", "") & vbCrLf & vbCrLf & _
               "You can undo this with Ctrl+Z." & vbCrLf & _
               "To clear only the colors : select the range, then Ctrl+Shift+C", _
               vbInformation, "Find Duplicates by Column"
    End If
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not find duplicates." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Find Duplicates by Column"
End Sub

'------------------------------------------------------------
' add the duplicate group-number / count columns  (from the Ctrl+Shift+V popup)
'   only rows whose picked-column combination appears twice or more get a number and a count.
'   numbered 1, 2, 3... in order of first appearance.  rows that are not duplicates stay blank.
'   two columns are inserted to the right of the table, and Ctrl+Z removes them.
'------------------------------------------------------------
Public Sub AddDuplicateGroupNo()
    Dim ws As Worksheet
    Dim data As Range
    Dim dic As Object, groupNo As Object
    Dim arr As Variant
    Dim visibleRange As Range, area As Range
    Dim visible() As Boolean
    Dim outArr() As Variant
    Dim key As String
    Dim r As Long, c As Long, startIdx As Long
    Dim rowCnt As Long, colCnt As Long
    Dim visibleCount As Long
    Dim groupCount As Long, dupRows As Long
    Dim pickedCount As Long
    Dim rightCol As Long, headerRow As Long
    Dim pickedNames As String
    Dim oldScreen As Boolean, oldCalc As Long
    Dim locked As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Len(mColAddr) = 0 Then
        MsgBox "Press Ctrl+Shift+V first to pick the key columns.", _
               vbInformation, "Duplicate Group Number"
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set data = ws.Range(mColAddr)
    rowCnt = data.Rows.Count
    colCnt = data.Columns.Count

    For c = 1 To colCnt
        If mColPick(c) Then pickedCount = pickedCount + 1
    Next c
    If pickedCount < 1 Then
        MsgBox "Please pick at least one key column.", vbExclamation, "Duplicate Group Number"
        Exit Sub
    End If

    '--- work out the visible rows (header row left out) ---
    ReDim visible(1 To rowCnt)
    On Error Resume Next
    Set visibleRange = data.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo Fail

    If visibleRange Is Nothing Then
        For r = 2 To rowCnt
            visible(r) = True
            visibleCount = visibleCount + 1
        Next r
    Else
        For Each area In visibleRange.Areas
            startIdx = area.Row - data.Row + 1
            For r = startIdx To startIdx + area.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not visible(r) Then
                        visible(r) = True
                        visibleCount = visibleCount + 1
                    End If
                End If
            Next r
        Next area
    End If

    If visibleCount < 2 Then
        MsgBox "Fewer than two visible data rows.", vbExclamation, "Duplicate Group Number"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    locked = True
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    arr = data.Value

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set groupNo = CreateObject("Scripting.Dictionary")
    groupNo.CompareMode = 1

    ' pass 1 : count by the picked column combination
    For r = 2 To rowCnt
      If visible(r) Then
        key = ""
        For c = 1 To colCnt
            If mColPick(c) Then key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic.Exists(key) Then
            dic(key) = dic(key) + 1
        Else
            dic.Add key, 1
        End If
      End If
    Next r

    ' pass 2 : number only the duplicates in order of first appearance and fill the output array
    ReDim outArr(1 To rowCnt - 1, 1 To 2)
    For r = 2 To rowCnt
      If visible(r) Then
        key = ""
        For c = 1 To colCnt
            If mColPick(c) Then key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            If Not groupNo.Exists(key) Then
                groupCount = groupCount + 1
                groupNo.Add key, groupCount
            End If
            outArr(r - 1, 1) = groupNo(key)
            outArr(r - 1, 2) = dic(key)
            dupRows = dupRows + 1
        End If
      End If
    Next r

    If dupRows = 0 Then
        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreen
        locked = False
        For c = 1 To colCnt
            If mColPick(c) Then
                If Len(pickedNames) > 0 Then pickedNames = pickedNames & " + "
                pickedNames = pickedNames & Trim$(CStr(data.Cells(1, c).Value))
            End If
        Next c
        If Not mQuiet Then
            MsgBox "No duplicate rows found." & vbCrLf & vbCrLf & _
                   "Key columns : " & pickedNames & vbCrLf & _
                   "Checked range : " & data.Address(False, False), _
                   vbInformation, "Duplicate Group Number"
        End If
        Exit Sub
    End If

    '--- insert two columns to the right of the table ---
    headerRow = data.Row
    rightCol = data.Column + colCnt
    ws.Range(ws.Columns(rightCol), ws.Columns(rightCol + 1)).Insert Shift:=xlToRight

    mInsWb = ws.Parent.Name
    mInsWs = ws.Name
    mInsCol = rightCol
    mInsCount = 2
    mInsOn = True

    ' an inserted column inherits the format on its left, so it is reset to General
    With ws.Range(ws.Cells(headerRow, rightCol), ws.Cells(headerRow + rowCnt - 1, rightCol + 1))
        .ClearFormats
        .NumberFormat = "General"
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(headerRow, rightCol), ws.Cells(headerRow, rightCol + 1))
        .Value = Array("DupGroup", "DupCount")
        .Font.Bold = True
    End With

    ws.Cells(headerRow + 1, rightCol).Resize(rowCnt - 1, 2).Value = outArr

    ws.Range(ws.Columns(rightCol), ws.Columns(rightCol + 1)).AutoFit

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    locked = False

    On Error Resume Next
    Application.OnUndo "Undo Duplicate Group Numbers", MacroPrefix() & "UndoInsertedColumns"
    Err.Clear
    On Error GoTo Fail

    For c = 1 To colCnt
        If mColPick(c) Then
            If Len(pickedNames) > 0 Then pickedNames = pickedNames & " + "
            pickedNames = pickedNames & Trim$(CStr(data.Cells(1, c).Value))
        End If
    Next c

    mDupGrp = groupCount
    mDupHit = dupRows
    If mQuiet Then Exit Sub

    MsgBox "Duplicates " & Format$(groupCount, "#,##0") & " kinds / " & _
           Format$(dupRows, "#,##0") & " rows were numbered." & vbCrLf & vbCrLf & _
           "Key columns : " & pickedNames & vbCrLf & _
           "Columns added : " & ColumnLetter(rightCol) & "(DupGroup) · " & _
                          ColumnLetter(rightCol + 1) & "(DupCount)" & vbCrLf & _
           "Checked range : " & data.Address(False, False) & _
           IIf(visibleCount < rowCnt - 1, "  (visible " & Format$(visibleCount, "#,##0") & " rows only)", "") & vbCrLf & vbCrLf & _
           "Rows that are not duplicates are left blank." & vbCrLf & _
           "Ctrl+Z removes the two columns that were added.", _
           vbInformation, "Duplicate Group Number"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    If locked Then
        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreen
    End If
    MsgBox "Could not add the numbers." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Duplicate Group Number"
End Sub

'------------------------------------------------------------
' highlight + group numbers in one go  (from the Ctrl+Shift+V popup)
'   calls the two existing features in turn. no code is copied.
'   the colors live in mDup* and the inserted columns in mIns*, so neither overwrites the other.
'   only the last OnUndo registration survives, so after calling both
'   a combined undo is registered once more, so one Ctrl+Z undoes both.
'------------------------------------------------------------
Public Sub HighlightAndNumberDuplicates()
    Dim pickedCount As Long
    Dim c As Long

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Len(mColAddr) = 0 Then
        MsgBox "Press Ctrl+Shift+V first to pick the key columns.", _
               vbInformation, "Highlight + Group Number"
        Exit Sub
    End If

    On Error Resume Next
    For c = LBound(mColPick) To UBound(mColPick)
        If mColPick(c) Then pickedCount = pickedCount + 1
    Next c
    Err.Clear
    On Error GoTo 0

    If pickedCount < 1 Then
        MsgBox "Please pick at least one key column.", vbExclamation, "Highlight + Group Number"
        Exit Sub
    End If

    ' cleared first, so it can tell whether this call added the columns
    mInsOn = False
    mDupGrp = 0
    mDupHit = 0

    On Error GoTo Fail
    mQuiet = True
    RunColumnDuplicates
    AddDuplicateGroupNo
    mQuiet = False
    On Error GoTo 0

    If Not mInsOn Then
        MsgBox "No duplicate rows found.", vbInformation, "Highlight + Group Number"
        Exit Sub
    End If

    On Error Resume Next
    Application.OnUndo "Undo Highlight + Group Numbers", MacroPrefix() & "UndoHighlightAndNumber"
    Err.Clear
    On Error GoTo 0

    MsgBox "Duplicates " & Format$(mDupGrp, "#,##0") & " kinds / " & _
           Format$(mDupHit, "#,##0") & " rows were highlighted and numbered." & vbCrLf & vbCrLf & _
           "The group number and the color match each other." & vbCrLf & _
           "When there are many groups the colors repeat, so use the numbers to tell them apart." & vbCrLf & vbCrLf & _
           "Ctrl+Z undoes both the colors and the two columns.", _
           vbInformation, "Highlight + Group Number"
    Exit Sub

Fail:
    mQuiet = False
    MsgBox "Could not complete it." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Highlight + Group Number"
End Sub

'--- undo : remove the two columns and restore the colors ---
Public Sub UndoHighlightAndNumber()
    On Error Resume Next
    UndoInsertedColumns
    UndoDuplicateColors
    Err.Clear
    On Error GoTo 0
    Application.StatusBar = "The highlighting and group numbers were undone."
    ScheduleStatusClear
End Sub

'--- column toggle wrappers (CommandBar OnAction cannot pass arguments, so there is one per number) ---
Public Sub ColPickToggle01()
    ColPickToggle 1
End Sub
Public Sub ColPickToggle02()
    ColPickToggle 2
End Sub
Public Sub ColPickToggle03()
    ColPickToggle 3
End Sub
Public Sub ColPickToggle04()
    ColPickToggle 4
End Sub
Public Sub ColPickToggle05()
    ColPickToggle 5
End Sub
Public Sub ColPickToggle06()
    ColPickToggle 6
End Sub
Public Sub ColPickToggle07()
    ColPickToggle 7
End Sub
Public Sub ColPickToggle08()
    ColPickToggle 8
End Sub
Public Sub ColPickToggle09()
    ColPickToggle 9
End Sub
Public Sub ColPickToggle10()
    ColPickToggle 10
End Sub
Public Sub ColPickToggle11()
    ColPickToggle 11
End Sub
Public Sub ColPickToggle12()
    ColPickToggle 12
End Sub
Public Sub ColPickToggle13()
    ColPickToggle 13
End Sub
Public Sub ColPickToggle14()
    ColPickToggle 14
End Sub
Public Sub ColPickToggle15()
    ColPickToggle 15
End Sub
Public Sub ColPickToggle16()
    ColPickToggle 16
End Sub
Public Sub ColPickToggle17()
    ColPickToggle 17
End Sub
Public Sub ColPickToggle18()
    ColPickToggle 18
End Sub
Public Sub ColPickToggle19()
    ColPickToggle 19
End Sub
Public Sub ColPickToggle20()
    ColPickToggle 20
End Sub
Public Sub ColPickToggle21()
    ColPickToggle 21
End Sub
Public Sub ColPickToggle22()
    ColPickToggle 22
End Sub
Public Sub ColPickToggle23()
    ColPickToggle 23
End Sub
Public Sub ColPickToggle24()
    ColPickToggle 24
End Sub
Public Sub ColPickToggle25()
    ColPickToggle 25
End Sub
Public Sub ColPickToggle26()
    ColPickToggle 26
End Sub
Public Sub ColPickToggle27()
    ColPickToggle 27
End Sub
Public Sub ColPickToggle28()
    ColPickToggle 28
End Sub
Public Sub ColPickToggle29()
    ColPickToggle 29
End Sub
Public Sub ColPickToggle30()
    ColPickToggle 30
End Sub
Public Sub ColPickToggle31()
    ColPickToggle 31
End Sub
Public Sub ColPickToggle32()
    ColPickToggle 32
End Sub
Public Sub ColPickToggle33()
    ColPickToggle 33
End Sub
Public Sub ColPickToggle34()
    ColPickToggle 34
End Sub
Public Sub ColPickToggle35()
    ColPickToggle 35
End Sub
Public Sub ColPickToggle36()
    ColPickToggle 36
End Sub
Public Sub ColPickToggle37()
    ColPickToggle 37
End Sub
Public Sub ColPickToggle38()
    ColPickToggle 38
End Sub
Public Sub ColPickToggle39()
    ColPickToggle 39
End Sub
Public Sub ColPickToggle40()
    ColPickToggle 40
End Sub
Public Sub ColPickToggle41()
    ColPickToggle 41
End Sub
Public Sub ColPickToggle42()
    ColPickToggle 42
End Sub
Public Sub ColPickToggle43()
    ColPickToggle 43
End Sub
Public Sub ColPickToggle44()
    ColPickToggle 44
End Sub
Public Sub ColPickToggle45()
    ColPickToggle 45
End Sub
Public Sub ColPickToggle46()
    ColPickToggle 46
End Sub
Public Sub ColPickToggle47()
    ColPickToggle 47
End Sub
Public Sub ColPickToggle48()
    ColPickToggle 48
End Sub
Public Sub ColPickToggle49()
    ColPickToggle 49
End Sub
Public Sub ColPickToggle50()
    ColPickToggle 50
End Sub
Public Sub ColPickToggle51()
    ColPickToggle 51
End Sub
Public Sub ColPickToggle52()
    ColPickToggle 52
End Sub
Public Sub ColPickToggle53()
    ColPickToggle 53
End Sub
Public Sub ColPickToggle54()
    ColPickToggle 54
End Sub
Public Sub ColPickToggle55()
    ColPickToggle 55
End Sub
Public Sub ColPickToggle56()
    ColPickToggle 56
End Sub
Public Sub ColPickToggle57()
    ColPickToggle 57
End Sub
Public Sub ColPickToggle58()
    ColPickToggle 58
End Sub
Public Sub ColPickToggle59()
    ColPickToggle 59
End Sub
Public Sub ColPickToggle60()
    ColPickToggle 60
End Sub

'--- 20-color palette for marking duplicates (10 light + 10 a shade deeper) ---
Private Function ColorPalette() As Variant
    ColorPalette = Array( _
        RGB(255, 235, 156), RGB(198, 239, 206), RGB(189, 215, 238), _
        RGB(255, 199, 206), RGB(226, 214, 244), RGB(255, 224, 178), _
        RGB(200, 232, 232), RGB(233, 218, 200), RGB(212, 230, 180), _
        RGB(244, 204, 228), _
        RGB(230, 200, 110), RGB(160, 210, 175), RGB(150, 185, 220), _
        RGB(205, 190, 235), RGB(225, 195, 130), RGB(150, 205, 205), _
        RGB(195, 180, 155), RGB(180, 205, 140), RGB(170, 200, 235), _
        RGB(200, 215, 175))
End Function

'============================================================
'  fill the formula down (Ctrl+Shift+1)
'============================================================
Public Sub FillFormulaDown()
    Dim ws As Worksheet
    Dim sel As Range
    Dim blk As Range
    Dim c As Range
    Dim endRow As Long
    Dim filledCount As Long
    Dim overwriteRows As Long
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Fill Formula Down"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell holding a formula, then run again.", vbExclamation, "Fill Formula Down"
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "Separate ranges cannot be processed at once.", _
               vbExclamation, "Fill Formula Down"
        Exit Sub
    End If
    If sel.Rows.Count > 1 Then
        MsgBox "Please select a single row." & vbCrLf & _
               "Select several cells across and each column is filled downward.", _
               vbExclamation, "Fill Formula Down"
        Exit Sub
    End If

    '--- find the last row of the table ---
    Set blk = SkipNonHeaderRows(sel.Cells(1, 1).CurrentRegion)
    If blk Is Nothing Then
        MsgBox "Could not find a table.", vbExclamation, "Fill Formula Down"
        Exit Sub
    End If
    endRow = blk.Row + blk.Rows.Count - 1

    If endRow <= sel.Row Then
        MsgBox "There are no rows below to fill." & vbCrLf & _
               "The cursor is on the last row, or no table was found.", _
               vbExclamation, "Fill Formula Down"
        Exit Sub
    End If

    '--- confirm when there is already data below ---
    overwriteRows = 0
    For Each c In sel.Cells
        overwriteRows = overwriteRows + Application.CountA(ws.Range(ws.Cells(sel.Row + 1, c.Column), _
                                                          ws.Cells(endRow, c.Column)))
    Next c

    If overwriteRows > 0 Then
        If MsgBox("There are already values below in " & Format$(overwriteRows, "#,##0") & " cells." & vbCrLf & _
                  "Overwrite?" & vbCrLf & vbCrLf & _
                  "Fill range : " & (sel.Row + 1) & " to row " & endRow & " rows" & vbCrLf & _
                  "You can undo this with Ctrl+Z.", _
                  vbOKCancel + vbExclamation, "Fill Formula Down") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    ' kept so it can be undone (the whole area being filled)
    Dim undoFlag As Range
    Set undoFlag = ws.Range(ws.Cells(sel.Row, sel.Column), _
                        ws.Cells(endRow, sel.Column + sel.Columns.Count - 1))
    SaveCriteriaBlock undoFlag

    ' fills only formulas and values, leaving formats alone (FillDown copies formats too)
    For Each c In sel.Cells
        ws.Range(ws.Cells(sel.Row, c.Column), ws.Cells(endRow, c.Column)).Formula = c.Formula
        filledCount = filledCount + 1
    Next c

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "Undo Fill Formula Down", MacroPrefix() & "UndoCriteriaBlock"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "Column " & filledCount & "" & (sel.Row + 1) & " to row " & endRow & " rows filled." & _
                            IIf(mCondOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not fill." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Fill Formula Down"
End Sub

'============================================================
'  insert an index column (Ctrl+Shift+2)
'============================================================
Public Sub InsertIndexColumn()
    Dim ws As Worksheet
    Dim blk As Range
    Dim headerRow As Long, endRow As Long, leftCol As Long
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Insert Index Column"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then Exit Sub

    Set blk = SkipNonHeaderRows(ActiveCell.CurrentRegion)
    If blk Is Nothing Then
        MsgBox "Could not find a table.", vbExclamation, "Insert Index Column"
        Exit Sub
    End If
    If blk.Rows.Count < 2 Then
        MsgBox "A header row and data rows are both required.", vbExclamation, "Insert Index Column"
        Exit Sub
    End If

    headerRow = blk.Row
    endRow = blk.Row + blk.Rows.Count - 1
    leftCol = blk.Column

    If MsgBox("This adds an index column at the far left of the table." & vbCrLf & vbCrLf & _
              "Location : " & Split(ws.Cells(1, leftCol).Address(True, False), "$")(0) & "Insert before column" & vbCrLf & _
              "Header : No" & vbCrLf & _
              "Formula : =ROW()-" & headerRow & vbCrLf & _
              "Range : " & (headerRow + 1) & " to row " & endRow & " rows" & vbCrLf & vbCrLf & _
              "You can undo this with Ctrl+Z.", _
              vbOKCancel + vbQuestion, "Insert Index Column") = vbCancel Then Exit Sub

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    ws.Columns(leftCol).Insert Shift:=xlToRight

    mInsWb = ws.Parent.Name
    mInsWs = ws.Name
    mInsCol = leftCol
    mInsCount = 1
    mInsOn = True

    With ws.Cells(headerRow, leftCol)
        .Value = "No"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(headerRow + 1, leftCol), ws.Cells(endRow, leftCol))
        .Formula = "=ROW()-" & headerRow
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(leftCol).AutoFit

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    Application.OnUndo "Undo Insert Index Column", MacroPrefix() & "UndoInsertedColumns"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "Index column added. (" & (endRow - headerRow) & " rows)" & _
                            "  |  Ctrl+Z to undo"
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not add the index column." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Insert Index Column"
End Sub

'--- undo the inserted column ---
Public Sub UndoInsertedColumns()
    Dim ws As Worksheet
    If Not mInsOn Or mInsCount < 1 Then
        MsgBox "There are no inserted columns to undo.", vbInformation, "Work Tools"
        Exit Sub
    End If
    On Error Resume Next
    Set ws = Application.Workbooks(mInsWb).Worksheets(mInsWs)
    ws.Range(ws.Columns(mInsCol), ws.Columns(mInsCol + mInsCount - 1)).Delete Shift:=xlToLeft
    mInsOn = False
    mInsCount = 0
    Err.Clear
    On Error GoTo 0
    Application.StatusBar = "The inserted columns were removed."
    ScheduleStatusClear
End Sub

'============================================================
'  find by value (Ctrl+Alt+F)
'    Ctrl+F looks in [Formulas] by default, so calculated results are missed.
'    this one searches by [Values] and selects every cell it finds.
'============================================================
Public Sub FindByValue()
    Dim ws As Worksheet
    Dim rng2 As Range
    Dim foundRange As Range
    Dim c As Range
    Dim firstAddr As String
    Dim findVal As String
    Dim cnt As Long
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Find by Value"
        Exit Sub
    End If
    Set ws = ActiveSheet

    ' within the selection when two or more cells are picked, otherwise across the sheet
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.CountLarge >= 2 Then
            Set rng2 = Selection
        Else
            Set rng2 = ws.UsedRange
        End If
    Else
        Set rng2 = ws.UsedRange
    End If

    findVal = InputBox("Value to find  (matches the result, selects every hit)" & vbCrLf & _
                      "After finding them, press Ctrl+Shift+D to check for duplicates" & vbCrLf & vbCrLf & _
                      "Range : " & rng2.Address(False, False), _
                      "Find by Value")

    If StrPtr(findVal) = 0 Then Exit Sub
    If Len(findVal) = 0 Then Exit Sub

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set c = rng2.Find(What:=findVal, LookIn:=xlValues, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False)

    If c Is Nothing Then
        Application.ScreenUpdating = oldScreen
        MsgBox "'" & findVal & "' was not found." & vbCrLf & vbCrLf & _
               "Searched range : " & rng2.Address(False, False), _
               vbInformation, "Find by Value"
        Exit Sub
    End If

    firstAddr = c.Address
    Do
        If foundRange Is Nothing Then
            Set foundRange = c
        Else
            Set foundRange = Application.Union(foundRange, c)
        End If
        cnt = cnt + 1
        If cnt >= 1000 Then Exit Do
        Set c = rng2.FindNext(c)
        If c Is Nothing Then Exit Do
    Loop While c.Address <> firstAddr

    foundRange.Select

    Application.ScreenUpdating = oldScreen

    Application.StatusBar = "'" & findVal & "' " & Format$(cnt, "#,##0") & " cells selected." & _
                            IIf(cnt >= 1000, "  (up to 1,000 cells)", "") & _
                            "   |  Ctrl+Shift+D to check duplicates among them"
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Not found." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Find by Value"
End Sub

'============================================================
'  multi VLOOKUP (Ctrl+Shift+L)
'    VLOOKUP only brings back the first match.
'    this one spreads every match out sideways.
'============================================================
Public Sub MultiLookup()
    Dim ws As Worksheet
    Dim data As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Multi VLOOKUP"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set data = PickColumnRange(ws)
    If data Is Nothing Then Exit Sub

    If StrComp(mLkWs, ws.Name, vbTextCompare) <> 0 Or _
       StrComp(mLkAddr, data.Address, vbTextCompare) <> 0 Then
        ReDim mLkKey(1 To data.Columns.Count)
        ReDim mLkVal(1 To data.Columns.Count)
        mLkWs = ws.Name
        mLkAddr = data.Address
        mLkFormula = False
        mLkDedup = True                       ' default : each combination only once
    End If

    On Error Resume Next
    If UBound(mLkKey) <> data.Columns.Count Then
        ReDim mLkKey(1 To data.Columns.Count)
        ReDim mLkVal(1 To data.Columns.Count)
    End If
    Err.Clear
    On Error GoTo 0

    MultiLookupMenu data
End Sub

Private Sub MultiLookupMenu(ByVal data As Range)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim i As Long
    Dim keyCount2 As Long, valCount2 As Long
    Dim name2 As String

    On Error Resume Next
    Application.CommandBars(BAR_MULTI_LOOKUP).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=BAR_MULTI_LOOKUP, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    For i = 1 To data.Columns.Count
        If mLkKey(i) Then keyCount2 = keyCount2 + 1
        If mLkVal(i) Then valCount2 = valCount2 + 1
    Next i

    AddMenuTitle cb, "Multi VLOOKUP  -  group matching values and spread them out"

    '--- (1) the key to group by ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "(1)  Group by  (more than one allowed)"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    For i = 1 To data.Columns.Count
        name2 = ColumnLabel(data, i)
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mLkKey(i), "     ■  ", "     □  ") & name2
        btn.OnAction = MacroPrefix() & "LkKeyToggle" & Format$(i, "00")
        btn.Style = msoButtonCaption
    Next i

    '--- (2) the values to bring back ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "(2)  Values to bring  (more than one allowed)"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    For i = 1 To data.Columns.Count
        name2 = ColumnLabel(data, i)
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mLkVal(i), "     ●  ", "     ○  ") & name2
        btn.OnAction = MacroPrefix() & "LkValToggle" & Format$(i, "00")
        btn.Style = msoButtonCaption
    Next i

    '--- (3) run ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Unique list to a new sheet   (one row per key)"
    btn.OnAction = MacroPrefix() & "MultiLookupToNewSheet"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True
    btn.Enabled = (keyCount2 > 0 And valCount2 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Spread beside the source   (same row count)"
    btn.OnAction = MacroPrefix() & "RunMultiLookup"
    btn.Style = msoButtonCaption
    btn.Enabled = (keyCount2 > 0 And valCount2 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Same values : " & IIf(mLkDedup, "each once  ->  duplicates included", "duplicates included  ->  each once")
    btn.OnAction = MacroPrefix() & "ToggleLookupDedup"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     When spreading beside : " & IIf(mLkFormula, "as formulas (updates itself)", "as values (fast)")
    btn.OnAction = MacroPrefix() & "ToggleLookupMode"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Clear all      (keys " & keyCount2 & " / value " & valCount2 & ")"
    btn.OnAction = MacroPrefix() & "LookupClearAll"
    btn.Style = msoButtonCaption

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "Could not show the menu." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Multi VLOOKUP"
End Sub

Private Function ColumnLabel(ByVal rg As Range, ByVal i As Long) As String
    Dim s As String
    On Error Resume Next
    s = Trim$(CStr(rg.Cells(1, i).Value))
    Err.Clear
    On Error GoTo 0
    If Len(s) = 0 Then s = "(no header) " & Split(rg.Cells(1, i).Address(True, False), "$")(1)
    ColumnLabel = s
End Function

Public Sub ToggleLookupMode()
    mLkFormula = Not mLkFormula
    MultiLookup
End Sub

Public Sub ToggleLookupDedup()
    mLkDedup = Not mLkDedup
    MultiLookup
End Sub

Public Sub LookupClearAll()
    Dim i As Long
    On Error Resume Next
    For i = LBound(mLkKey) To UBound(mLkKey)
        mLkKey(i) = False
        mLkVal(i) = False
    Next i
    Err.Clear
    On Error GoTo 0
    MultiLookup
End Sub

Private Sub LkKeyToggle(ByVal no As Long)
    On Error Resume Next
    mLkKey(no) = Not mLkKey(no)
    If mLkKey(no) Then mLkVal(no) = False
    Err.Clear
    On Error GoTo 0
    MultiLookup
End Sub

Private Sub LkValToggle(ByVal no As Long)
    On Error Resume Next
    mLkVal(no) = Not mLkVal(no)
    If mLkVal(no) Then mLkKey(no) = False
    Err.Clear
    On Error GoTo 0
    MultiLookup
End Sub

'--- the lookup itself ---
Public Sub RunMultiLookup()
    Dim ws As Worksheet
    Dim data As Range
    Dim dic As Object
    Dim arr As Variant
    Dim key As String
    Dim r As Long, c As Long, i As Long, j As Long
    Dim rowCnt As Long, colCnt As Long
    Dim maxPerKey As Long
    Dim valCols() As Long, valCount As Long
    Dim pasteCol As Long, needCols As Long
    Dim hasRight As Boolean
    Dim inserted As Boolean
    Dim header As String
    Dim oldScreen As Boolean, oldCalc As Long
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = ActiveSheet
    Set data = ws.Range(mLkAddr)
    rowCnt = data.Rows.Count
    colCnt = data.Columns.Count

    If rowCnt > 20000 Then
        If MsgBox(Format$(rowCnt - 1, "#,##0") & " rows will be processed. This may take a while." & vbCrLf & _
                  "Do you want to continue?", vbOKCancel + vbQuestion, "Multi VLOOKUP") = vbCancel Then Exit Sub
    End If

    ' the list of value columns
    ReDim valCols(1 To colCnt)
    For c = 1 To colCnt
        If mLkVal(c) Then
            valCount = valCount + 1
            valCols(valCount) = c
        End If
    Next c

    arr = data.Value

    '--- count per key to find the largest (with duplicates removed where asked) ---
    Dim combo2 As Object
    Dim vkey2 As String

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set combo2 = CreateObject("Scripting.Dictionary")
    combo2.CompareMode = 1

    For r = 2 To rowCnt
        key = LookupKey(arr, r, colCnt)

        Dim cellOk As Boolean
        cellOk = True
        If mLkDedup Then
            vkey2 = key & Chr$(2) & ValueKey(arr, r, valCols, valCount)
            If combo2.Exists(vkey2) Then
                cellOk = False
            Else
                combo2.Add vkey2, 1
            End If
        End If

        If cellOk Then
            If dic.Exists(key) Then
                dic(key) = dic(key) + 1
            Else
                dic.Add key, 1
            End If
            If dic(key) > maxPerKey Then maxPerKey = dic(key)
        End If
    Next r

    If maxPerKey < 1 Then
        MsgBox "There is no data to process.", vbExclamation, "Multi VLOOKUP"
        Exit Sub
    End If

    needCols = IIf(mLkDedup, 2, 1) + maxPerKey * valCount   ' the count column + (hits x value columns)

    If needCols > 200 Then
        MsgBox "Columns needed: " & needCols & ", which is too many." & vbCrLf & _
               "Up to " & maxPerKey & " rows." & vbCrLf & _
               "Please pick fewer value columns.", vbExclamation, "Multi VLOOKUP"
        Exit Sub
    End If

    pasteCol = data.Column + colCnt

    ' check whether there is data to the right
    On Error Resume Next
    hasRight = (Application.CountA(ws.Range(ws.Cells(data.Row, pasteCol), _
                                              ws.Cells(data.Row + rowCnt - 1, pasteCol + needCols - 1))) > 0)
    Err.Clear
    On Error GoTo Fail

    If mLkFormula Then
        If MsgBox("Formulas make the file heavier and can slow Excel down." & vbCrLf & _
                  "Saving the file first is recommended." & vbCrLf & vbCrLf & _
                  "Do you want to continue?", vbOKCancel + vbExclamation, "Multi VLOOKUP") = vbCancel Then Exit Sub
    End If

    If hasRight Then
        If MsgBox("There is already data to the right." & vbCrLf & _
                  " empty column" & needCols & "Insert that many columns and fill?" & vbCrLf & vbCrLf & _
                  "Existing data will shift to the right." & vbCrLf & _
                  "You can undo this with Ctrl+Z.", _
                  vbOKCancel + vbQuestion, "Multi VLOOKUP") = vbCancel Then Exit Sub
        inserted = True
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If inserted Then
        ws.Range(ws.Columns(pasteCol), ws.Columns(pasteCol + needCols - 1)).Insert Shift:=xlToRight
        mInsWb = ws.Parent.Name
        mInsWs = ws.Name
        mInsCol = pasteCol
        mInsCount = needCols
        mInsOn = True
    Else
        SaveCriteriaBlock ws.Range(ws.Cells(data.Row, pasteCol), _
                            ws.Cells(data.Row + rowCnt - 1, pasteCol + needCols - 1))
        mInsOn = False
    End If

    '--- the headers ---
    j = pasteCol - 1
    If mLkDedup Then
        j = j + 1
        With ws.Cells(data.Row, j)
            .Value = "TotalCount"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        j = j + 1
        With ws.Cells(data.Row, j)
            .Value = "UniqueCount"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Else
        j = j + 1
        With ws.Cells(data.Row, j)
            .Value = "Count"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If

    For c = 1 To valCount
        header = Trim$(CStr(data.Cells(1, valCols(c)).Value))
        If Len(header) = 0 Then header = "Value"
        For i = 1 To maxPerKey
            j = j + 1
            With ws.Cells(data.Row, j)
                .Value = header & i
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
        Next i
    Next c

    ' apply the value columns' number formats up front
    j = pasteCol + IIf(mLkDedup, 2, 1) - 1
    For c = 1 To valCount
        For i = 1 To maxPerKey
            j = j + 1
            ws.Cells(data.Row + 1, j).Resize(rowCnt - 1, 1).NumberFormat = _
                data.Cells(2, valCols(c)).NumberFormat
        Next i
    Next c

    '--- fill the body ---
    FillMultiLookup ws, data, arr, rowCnt, colCnt, pasteCol, maxPerKey, valCols, valCount

    ws.Range(ws.Columns(pasteCol), ws.Columns(pasteCol + needCols - 1)).AutoFit

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    If inserted Then
        Application.OnUndo "Undo Multi VLOOKUP", MacroPrefix() & "UndoInsertedColumns"
    Else
        Application.OnUndo "Undo Multi VLOOKUP", MacroPrefix() & "UndoCriteriaBlock"
    End If
    Err.Clear
    On Error GoTo Fail

    MsgBox "Lookup finished." & vbCrLf & vbCrLf & _
           "Largest group has " & maxPerKey & "" & _
           IIf(mLkDedup, "  (each combination once)", "  (duplicates included)") & vbCrLf & _
           "Columns created : " & needCols & "" & vbCrLf & _
           "Inserted as : " & IIf(mLkFormula, "formulas", "Value") & vbCrLf & vbCrLf & _
           "You can undo this with Ctrl+Z.", vbInformation, "Multi VLOOKUP"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
    MsgBox "Could not run the lookup." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Multi VLOOKUP"
End Sub

'--- the key string for one row ---
Private Function LookupKey(ByVal arr As Variant, ByVal r As Long, ByVal colCnt As Long) As String
    Dim c As Long
    Dim s As String
    For c = 1 To colCnt
        If mLkKey(c) Then s = s & CStr(arr(r, c)) & Chr$(1)
    Next c
    LookupKey = s
End Function

'--- fill the body ---
Private Sub FillMultiLookup(ByVal ws As Worksheet, ByVal data As Range, _
                           ByVal arr As Variant, ByVal rowCnt As Long, ByVal colCnt As Long, _
                           ByVal pasteCol As Long, ByVal maxPerKey As Long, _
                           ByRef valCols() As Long, ByVal valCount As Long)
    Dim list2 As Object, combo As Object, totalCount As Object
    Dim key As String, vkey As String
    Dim r As Long, c As Long, i As Long, j As Long
    Dim rowList As Variant
    Dim outArr() As Variant
    Dim colLetter() As String
    Dim firstRow As Long
    Dim countCols As Long

    Set list2 = CreateObject("Scripting.Dictionary")
    list2.CompareMode = 1
    Set combo = CreateObject("Scripting.Dictionary")
    combo.CompareMode = 1
    Set totalCount = CreateObject("Scripting.Dictionary")
    totalCount.CompareMode = 1

    countCols = IIf(mLkDedup, 2, 1)

    For r = 2 To rowCnt
        key = LookupKey(arr, r, colCnt)

        If totalCount.Exists(key) Then
            totalCount(key) = totalCount(key) + 1
        Else
            totalCount.Add key, 1
        End If

        Dim keepIt As Boolean
        keepIt = True
        If mLkDedup Then
            vkey = key & Chr$(2) & ValueKey(arr, r, valCols, valCount)
            If combo.Exists(vkey) Then
                keepIt = False
            Else
                combo.Add vkey, 1
            End If
        End If

        If keepIt Then
            If list2.Exists(key) Then
                list2(key) = list2(key) & "," & r
            Else
                list2.Add key, CStr(r)
            End If
        End If
    Next r

    firstRow = data.Row

    ' build the value columns' letters up front
    ReDim colLetter(1 To valCount)
    For c = 1 To valCount
        colLetter(c) = ColumnLetter(data.Column + valCols(c) - 1)
    Next c

    ReDim outArr(1 To rowCnt - 1, 1 To countCols + maxPerKey * valCount)

    For r = 2 To rowCnt
        key = LookupKey(arr, r, colCnt)
        rowList = Split(list2(key), ",")

        If mLkDedup Then
            outArr(r - 1, 1) = totalCount(key)
            outArr(r - 1, 2) = UBound(rowList) + 1
        Else
            outArr(r - 1, 1) = UBound(rowList) + 1
        End If

        j = countCols
        For c = 1 To valCount
            For i = 0 To maxPerKey - 1
                j = j + 1
                If i <= UBound(rowList) Then
                    If mLkFormula Then
                        outArr(r - 1, j) = "=" & colLetter(c) & (firstRow + CLng(rowList(i)) - 1)
                    Else
                        outArr(r - 1, j) = arr(CLng(rowList(i)), valCols(c))
                    End If
                End If
            Next i
        Next c
    Next r

    ws.Range(ws.Cells(firstRow + 1, pasteCol), _
             ws.Cells(firstRow + rowCnt - 1, pasteCol + countCols + maxPerKey * valCount - 1)).Formula = outArr
End Sub

'--- the combined value-column key (used to spot duplicates) ---
Private Function ValueKey(ByVal arr As Variant, ByVal r As Long, _
                      ByRef valCols() As Long, ByVal valCount As Long) As String
    Dim c As Long
    Dim s As String
    For c = 1 To valCount
        s = s & CStr(arr(r, valCols(c))) & Chr$(1)
    Next c
    ValueKey = s
End Function

'============================================================
'  select to the end of the table (Ctrl+Shift+3)
'============================================================
Public Sub SelectToEnd()
    Dim ws As Worksheet
    Dim sel As Range
    Dim blk As Range
    Dim endRow As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "Separate ranges cannot be processed at once.", _
               vbExclamation, "Select to End of Table"
        Exit Sub
    End If

    On Error GoTo Fail

    Set blk = SkipNonHeaderRows(sel.Cells(1, 1).CurrentRegion)
    If blk Is Nothing Then
        MsgBox "Could not find a table.", vbExclamation, "Select to End of Table"
        Exit Sub
    End If

    endRow = blk.Row + blk.Rows.Count - 1
    If endRow <= sel.Row + sel.Rows.Count - 1 Then
        Application.StatusBar = "The selection already reaches the end of the table. (" & endRow & " rows)"
        ScheduleStatusClear
        Exit Sub
    End If

    ws.Range(ws.Cells(sel.Row, sel.Column), _
             ws.Cells(endRow, sel.Column + sel.Columns.Count - 1)).Select

    Application.StatusBar = sel.Row & " to row " & endRow & "Select rows  (" & _
                            Format$(endRow - sel.Row + 1, "#,##0") & " rows x " & _
                            sel.Columns.Count & " columns)"
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "Could not select." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Select to End of Table"
End Sub

'============================================================
'  key columns - extract to a new sheet (duplicate rows / unique list)
'============================================================
Public Sub ExtractDuplicateRows()
    ExtractByColumns True
End Sub

Public Sub ExtractUniqueList()
    ExtractByColumns False
End Sub

Private Sub ExtractByColumns(ByVal dupOnly As Boolean)
    Dim ws As Worksheet, newSheet As Worksheet
    Dim data As Range
    Dim dic As Object, taken As Object
    Dim arr As Variant, outArr() As Variant
    Dim visibleRange As Range, area As Range
    Dim visible() As Boolean
    Dim key As String
    Dim r As Long, c As Long, i As Long, startIdx As Long
    Dim rowCnt As Long, colCnt As Long
    Dim visibleCount As Long
    Dim takenCount As Long, outCol As Long
    Dim pickedNames As String
    Dim name2 As String
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = ActiveSheet
    Set data = ws.Range(mColAddr)
    rowCnt = data.Rows.Count
    colCnt = data.Columns.Count

    ReDim visible(1 To rowCnt)
    On Error Resume Next
    Set visibleRange = data.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo 0

    If visibleRange Is Nothing Then
        For r = 2 To rowCnt
            visible(r) = True
            visibleCount = visibleCount + 1
        Next r
    Else
        For Each area In visibleRange.Areas
            startIdx = area.Row - data.Row + 1
            For r = startIdx To startIdx + area.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not visible(r) Then
                        visible(r) = True
                        visibleCount = visibleCount + 1
                    End If
                End If
            Next r
        Next area
    End If

    If visibleCount < 1 Then
        MsgBox "There are no visible data rows.", vbExclamation, "Extract by Columns"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = data.Value

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1

    For r = 2 To rowCnt
      If visible(r) Then
        key = GroupKey(arr, r, colCnt)
        If dic.Exists(key) Then
            dic(key) = dic(key) + 1
        Else
            dic.Add key, 1
        End If
      End If
    Next r

    outCol = 0
    For c = 1 To colCnt
        If mColPick(c) Then outCol = outCol + 1
    Next c

    If dupOnly Then
        ReDim outArr(1 To visibleCount + 1, 1 To colCnt)
        For c = 1 To colCnt
            outArr(1, c) = arr(1, c)
        Next c
        takenCount = 0
        For r = 2 To rowCnt
          If visible(r) Then
            key = GroupKey(arr, r, colCnt)
            If dic(key) >= 2 Then
                takenCount = takenCount + 1
                For c = 1 To colCnt
                    outArr(takenCount + 1, c) = arr(r, c)
                Next c
            End If
          End If
        Next r
    Else
        Set taken = CreateObject("Scripting.Dictionary")
        taken.CompareMode = 1
        ReDim outArr(1 To dic.Count + 1, 1 To outCol)
        i = 0
        For c = 1 To colCnt
            If mColPick(c) Then
                i = i + 1
                outArr(1, i) = arr(1, c)
            End If
        Next c
        takenCount = 0
        For r = 2 To rowCnt
          If visible(r) Then
            key = GroupKey(arr, r, colCnt)
            If Not taken.Exists(key) Then
                taken.Add key, 1
                takenCount = takenCount + 1
                i = 0
                For c = 1 To colCnt
                    If mColPick(c) Then
                        i = i + 1
                        outArr(takenCount + 1, i) = arr(r, c)
                    End If
                Next c
            End If
          End If
        Next r
    End If

    Application.ScreenUpdating = oldScreen

    If takenCount = 0 Then
        MsgBox IIf(dupOnly, "No duplicate rows found.", "There is no data to extract."), _
               vbInformation, "Extract by Columns"
        Exit Sub
    End If

    For c = 1 To colCnt
        If mColPick(c) Then
            If Len(pickedNames) > 0 Then pickedNames = pickedNames & " + "
            pickedNames = pickedNames & Trim$(CStr(data.Cells(1, c).Value))
        End If
    Next c

    Application.ScreenUpdating = False
    name2 = NewSheetName(ws.Parent, IIf(dupOnly, "Dup_", "Unique_") & ws.Name)
    Set newSheet = ws.Parent.Worksheets.Add(Before:=ws)
    On Error Resume Next
    newSheet.Name = name2
    Err.Clear
    On Error GoTo Fail

    With newSheet.Range("A1")
        .Value = IIf(dupOnly, "Extract duplicate rows", "Unique list") & _
                 "  |  source : " & ws.Name & "!" & data.Address(False, False) & _
                 "  |  key columns : " & pickedNames
        .Font.Size = 9
        .Font.Color = RGB(110, 110, 115)
    End With
    newSheet.Columns("A").ColumnWidth = 12

    ' move the number formats across before the values (so 00123 does not become 123)
    If dupOnly Then
        For c = 1 To colCnt
            newSheet.Cells(4, c).Resize(takenCount, 1).NumberFormat = data.Cells(2, c).NumberFormat
        Next c
    Else
        i = 0
        For c = 1 To colCnt
            If mColPick(c) Then
                i = i + 1
                newSheet.Cells(4, i).Resize(takenCount, 1).NumberFormat = data.Cells(2, c).NumberFormat
            End If
        Next c
    End If

    newSheet.Range("A3").Resize(takenCount + 1, UBound(outArr, 2)).Value = outArr
    With newSheet.Range("A3").Resize(1, UBound(outArr, 2))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    newSheet.Range("A3").Resize(takenCount + 1, UBound(outArr, 2)).Columns.AutoFit
    newSheet.Activate
    newSheet.Range("A4").Select

    Application.ScreenUpdating = oldScreen

    MsgBox IIf(dupOnly, "Duplicate rows ", "Unique list ") & Format$(takenCount, "#,##0") & " rows were extracted." & vbCrLf & vbCrLf & _
           "Sheet : " & name2 & vbCrLf & _
           "Key columns : " & pickedNames & vbCrLf & _
           "Source : " & ws.Name & "!" & data.Address(False, False) & _
           IIf(visibleCount < rowCnt - 1, "  (visible " & visibleCount & " rows only)", ""), _
           vbInformation, "Extract by Columns"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not extract it." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Extract by Columns"
End Sub

Private Function GroupKey(ByVal arr As Variant, ByVal r As Long, ByVal colCnt As Long) As String
    Dim c As Long
    Dim s As String
    For c = 1 To colCnt
        If mColPick(c) Then s = s & CStr(arr(r, c)) & Chr$(1)
    Next c
    GroupKey = s
End Function

Private Function NewSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim candidate As String, suffix As String
    Dim i As Long

    baseName = Replace(baseName, ":", "")
    baseName = Replace(baseName, "/", "")
    baseName = Replace(baseName, "?", "")
    baseName = Replace(baseName, "*", "")
    baseName = Replace(baseName, "[", "")
    baseName = Replace(baseName, "]", "")
    If Len(baseName) > 31 Then baseName = Left$(baseName, 31)

    If Not SheetExists(wb, baseName) Then
        NewSheetName = baseName
        Exit Function
    End If

    For i = 2 To 999
        suffix = CStr(i)
        candidate = baseName
        If Len(candidate) + Len(suffix) > 31 Then candidate = Left$(candidate, 31 - Len(suffix))
        candidate = candidate & suffix
        If Not SheetExists(wb, candidate) Then
            NewSheetName = candidate
            Exit Function
        End If
    Next i
    NewSheetName = "Picked_" & Format$(Now, "hhnnss")
End Function

'============================================================
'  multi VLOOKUP - unique list onto a new sheet
'============================================================
Public Sub MultiLookupToNewSheet()
    Dim ws As Worksheet, newSheet As Worksheet
    Dim data As Range
    Dim order2 As Object, bucket As Object, totalCount As Object, combo As Object
    Dim arr As Variant, outArr() As Variant
    Dim visibleRange As Range, area As Range
    Dim visible() As Boolean
    Dim valCols() As Long, keyCols() As Long
    Dim valCount As Long, keyCount As Long
    Dim key As String, vkey As String, name2 As String
    Dim r As Long, c As Long, i As Long, j As Long, k As Long, startIdx As Long
    Dim rowCnt As Long, colCnt As Long
    Dim visibleCount As Long
    Dim maxPerKey As Long, keyCount3 As Long
    Dim countCols As Long
    Dim keepIt As Boolean
    Dim rowList As Variant
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = ActiveSheet
    Set data = ws.Range(mLkAddr)
    rowCnt = data.Rows.Count
    colCnt = data.Columns.Count
    countCols = IIf(mLkDedup, 2, 1)

    ReDim keyCols(1 To colCnt)
    ReDim valCols(1 To colCnt)
    For c = 1 To colCnt
        If mLkKey(c) Then
            keyCount = keyCount + 1
            keyCols(keyCount) = c
        End If
        If mLkVal(c) Then
            valCount = valCount + 1
            valCols(valCount) = c
        End If
    Next c

    ReDim visible(1 To rowCnt)
    On Error Resume Next
    Set visibleRange = data.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo 0
    If visibleRange Is Nothing Then
        For r = 2 To rowCnt
            visible(r) = True
            visibleCount = visibleCount + 1
        Next r
    Else
        For Each area In visibleRange.Areas
            startIdx = area.Row - data.Row + 1
            For r = startIdx To startIdx + area.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not visible(r) Then
                        visible(r) = True
                        visibleCount = visibleCount + 1
                    End If
                End If
            Next r
        Next area
    End If

    If visibleCount < 1 Then
        MsgBox "There are no visible data rows.", vbExclamation, "Multi VLOOKUP"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = data.Value

    Set order2 = CreateObject("Scripting.Dictionary")
    order2.CompareMode = 1
    Set bucket = CreateObject("Scripting.Dictionary")
    bucket.CompareMode = 1
    Set totalCount = CreateObject("Scripting.Dictionary")
    totalCount.CompareMode = 1
    Set combo = CreateObject("Scripting.Dictionary")
    combo.CompareMode = 1

    For r = 2 To rowCnt
      If visible(r) Then
        key = LookupKey(arr, r, colCnt)

        If totalCount.Exists(key) Then
            totalCount(key) = totalCount(key) + 1
        Else
            totalCount.Add key, 1
        End If

        keepIt = True
        If mLkDedup Then
            vkey = key & Chr$(2) & ValueKey(arr, r, valCols, valCount)
            If combo.Exists(vkey) Then
                keepIt = False
            Else
                combo.Add vkey, 1
            End If
        End If

        If keepIt Then
            If bucket.Exists(key) Then
                bucket(key) = bucket(key) & "," & r
            Else
                bucket.Add key, CStr(r)
                keyCount3 = keyCount3 + 1
                order2.Add keyCount3, key
            End If
            i = UBound(Split(bucket(key), ",")) + 1
            If i > maxPerKey Then maxPerKey = i
        End If
      End If
    Next r

    If keyCount + countCols + maxPerKey * valCount > 200 Then
        Application.ScreenUpdating = oldScreen
        MsgBox "Too many columns are needed." & vbCrLf & _
               "Up to " & maxPerKey & " of them." & vbCrLf & _
               "Please pick fewer value columns.", vbExclamation, "Multi VLOOKUP"
        Exit Sub
    End If

    ReDim outArr(1 To keyCount3 + 1, 1 To keyCount + countCols + maxPerKey * valCount)

    j = 0
    For i = 1 To keyCount
        j = j + 1
        outArr(1, j) = arr(1, keyCols(i))
    Next i
    If mLkDedup Then
        j = j + 1
        outArr(1, j) = "TotalCount"
        j = j + 1
        outArr(1, j) = "UniqueCount"
    Else
        j = j + 1
        outArr(1, j) = "Count"
    End If
    For c = 1 To valCount
        name2 = Trim$(CStr(arr(1, valCols(c))))
        If Len(name2) = 0 Then name2 = "Value"
        For i = 1 To maxPerKey
            j = j + 1
            outArr(1, j) = name2 & i
        Next i
    Next c

    For k = 1 To keyCount3
        key = order2(k)
        rowList = Split(bucket(key), ",")
        r = CLng(rowList(0))

        j = 0
        For i = 1 To keyCount
            j = j + 1
            outArr(k + 1, j) = arr(r, keyCols(i))
        Next i
        If mLkDedup Then
            j = j + 1
            outArr(k + 1, j) = totalCount(key)
            j = j + 1
            outArr(k + 1, j) = UBound(rowList) + 1
        Else
            j = j + 1
            outArr(k + 1, j) = UBound(rowList) + 1
        End If

        For c = 1 To valCount
            For i = 1 To maxPerKey
                j = j + 1
                If i - 1 <= UBound(rowList) Then
                    outArr(k + 1, j) = arr(CLng(rowList(i - 1)), valCols(c))
                End If
            Next i
        Next c
    Next k

    name2 = NewSheetName(ws.Parent, "Group_" & ws.Name)
    Set newSheet = ws.Parent.Worksheets.Add(Before:=ws)
    On Error Resume Next
    newSheet.Name = name2
    Err.Clear
    On Error GoTo Fail

    With newSheet.Range("A1")
        .Value = "Multi VLOOKUP  |  source : " & ws.Name & "!" & data.Address(False, False) & _
                 "  |  keys " & keyCount & " / value " & valCount & _
                 IIf(mLkDedup, "  |  each combination once", "  |  duplicates included")
        .Font.Size = 9
        .Font.Color = RGB(110, 110, 115)
    End With
    newSheet.Columns("A").ColumnWidth = 12

    ' move the number formats across before the values
    j = 0
    For i = 1 To keyCount
        j = j + 1
        newSheet.Cells(4, j).Resize(keyCount3, 1).NumberFormat = data.Cells(2, keyCols(i)).NumberFormat
    Next i
    j = j + countCols                                  ' the count column stays General
    For c = 1 To valCount
        For i = 1 To maxPerKey
            j = j + 1
            newSheet.Cells(4, j).Resize(keyCount3, 1).NumberFormat = data.Cells(2, valCols(c)).NumberFormat
        Next i
    Next c

    newSheet.Range("A3").Resize(keyCount3 + 1, UBound(outArr, 2)).Value = outArr
    With newSheet.Range("A3").Resize(1, UBound(outArr, 2))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    newSheet.Range("A3").Resize(keyCount3 + 1, UBound(outArr, 2)).Columns.AutoFit
    newSheet.Activate
    newSheet.Range("A4").Select

    Application.ScreenUpdating = oldScreen

    MsgBox "Unique list " & Format$(keyCount3, "#,##0") & " rows created." & vbCrLf & vbCrLf & _
           "Sheet : " & name2 & vbCrLf & _
           "Largest group has " & maxPerKey & "" & _
           IIf(mLkDedup, "  (each combination once)", "  (duplicates included)") & vbCrLf & _
           "Source : " & ws.Name & "!" & data.Address(False, False) & _
           IIf(visibleCount < rowCnt - 1, "  (visible " & visibleCount & " rows only)", ""), _
           vbInformation, "Multi VLOOKUP"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not create it." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Multi VLOOKUP"
End Sub

'--- lookup key / value toggle wrappers ---
Public Sub LkKeyToggle01()
    LkKeyToggle 1
End Sub
Public Sub LkKeyToggle02()
    LkKeyToggle 2
End Sub
Public Sub LkKeyToggle03()
    LkKeyToggle 3
End Sub
Public Sub LkKeyToggle04()
    LkKeyToggle 4
End Sub
Public Sub LkKeyToggle05()
    LkKeyToggle 5
End Sub
Public Sub LkKeyToggle06()
    LkKeyToggle 6
End Sub
Public Sub LkKeyToggle07()
    LkKeyToggle 7
End Sub
Public Sub LkKeyToggle08()
    LkKeyToggle 8
End Sub
Public Sub LkKeyToggle09()
    LkKeyToggle 9
End Sub
Public Sub LkKeyToggle10()
    LkKeyToggle 10
End Sub
Public Sub LkKeyToggle11()
    LkKeyToggle 11
End Sub
Public Sub LkKeyToggle12()
    LkKeyToggle 12
End Sub
Public Sub LkKeyToggle13()
    LkKeyToggle 13
End Sub
Public Sub LkKeyToggle14()
    LkKeyToggle 14
End Sub
Public Sub LkKeyToggle15()
    LkKeyToggle 15
End Sub
Public Sub LkKeyToggle16()
    LkKeyToggle 16
End Sub
Public Sub LkKeyToggle17()
    LkKeyToggle 17
End Sub
Public Sub LkKeyToggle18()
    LkKeyToggle 18
End Sub
Public Sub LkKeyToggle19()
    LkKeyToggle 19
End Sub
Public Sub LkKeyToggle20()
    LkKeyToggle 20
End Sub
Public Sub LkKeyToggle21()
    LkKeyToggle 21
End Sub
Public Sub LkKeyToggle22()
    LkKeyToggle 22
End Sub
Public Sub LkKeyToggle23()
    LkKeyToggle 23
End Sub
Public Sub LkKeyToggle24()
    LkKeyToggle 24
End Sub
Public Sub LkKeyToggle25()
    LkKeyToggle 25
End Sub
Public Sub LkKeyToggle26()
    LkKeyToggle 26
End Sub
Public Sub LkKeyToggle27()
    LkKeyToggle 27
End Sub
Public Sub LkKeyToggle28()
    LkKeyToggle 28
End Sub
Public Sub LkKeyToggle29()
    LkKeyToggle 29
End Sub
Public Sub LkKeyToggle30()
    LkKeyToggle 30
End Sub
Public Sub LkKeyToggle31()
    LkKeyToggle 31
End Sub
Public Sub LkKeyToggle32()
    LkKeyToggle 32
End Sub
Public Sub LkKeyToggle33()
    LkKeyToggle 33
End Sub
Public Sub LkKeyToggle34()
    LkKeyToggle 34
End Sub
Public Sub LkKeyToggle35()
    LkKeyToggle 35
End Sub
Public Sub LkKeyToggle36()
    LkKeyToggle 36
End Sub
Public Sub LkKeyToggle37()
    LkKeyToggle 37
End Sub
Public Sub LkKeyToggle38()
    LkKeyToggle 38
End Sub
Public Sub LkKeyToggle39()
    LkKeyToggle 39
End Sub
Public Sub LkKeyToggle40()
    LkKeyToggle 40
End Sub
Public Sub LkKeyToggle41()
    LkKeyToggle 41
End Sub
Public Sub LkKeyToggle42()
    LkKeyToggle 42
End Sub
Public Sub LkKeyToggle43()
    LkKeyToggle 43
End Sub
Public Sub LkKeyToggle44()
    LkKeyToggle 44
End Sub
Public Sub LkKeyToggle45()
    LkKeyToggle 45
End Sub
Public Sub LkKeyToggle46()
    LkKeyToggle 46
End Sub
Public Sub LkKeyToggle47()
    LkKeyToggle 47
End Sub
Public Sub LkKeyToggle48()
    LkKeyToggle 48
End Sub
Public Sub LkKeyToggle49()
    LkKeyToggle 49
End Sub
Public Sub LkKeyToggle50()
    LkKeyToggle 50
End Sub
Public Sub LkKeyToggle51()
    LkKeyToggle 51
End Sub
Public Sub LkKeyToggle52()
    LkKeyToggle 52
End Sub
Public Sub LkKeyToggle53()
    LkKeyToggle 53
End Sub
Public Sub LkKeyToggle54()
    LkKeyToggle 54
End Sub
Public Sub LkKeyToggle55()
    LkKeyToggle 55
End Sub
Public Sub LkKeyToggle56()
    LkKeyToggle 56
End Sub
Public Sub LkKeyToggle57()
    LkKeyToggle 57
End Sub
Public Sub LkKeyToggle58()
    LkKeyToggle 58
End Sub
Public Sub LkKeyToggle59()
    LkKeyToggle 59
End Sub
Public Sub LkKeyToggle60()
    LkKeyToggle 60
End Sub
Public Sub LkValToggle01()
    LkValToggle 1
End Sub
Public Sub LkValToggle02()
    LkValToggle 2
End Sub
Public Sub LkValToggle03()
    LkValToggle 3
End Sub
Public Sub LkValToggle04()
    LkValToggle 4
End Sub
Public Sub LkValToggle05()
    LkValToggle 5
End Sub
Public Sub LkValToggle06()
    LkValToggle 6
End Sub
Public Sub LkValToggle07()
    LkValToggle 7
End Sub
Public Sub LkValToggle08()
    LkValToggle 8
End Sub
Public Sub LkValToggle09()
    LkValToggle 9
End Sub
Public Sub LkValToggle10()
    LkValToggle 10
End Sub
Public Sub LkValToggle11()
    LkValToggle 11
End Sub
Public Sub LkValToggle12()
    LkValToggle 12
End Sub
Public Sub LkValToggle13()
    LkValToggle 13
End Sub
Public Sub LkValToggle14()
    LkValToggle 14
End Sub
Public Sub LkValToggle15()
    LkValToggle 15
End Sub
Public Sub LkValToggle16()
    LkValToggle 16
End Sub
Public Sub LkValToggle17()
    LkValToggle 17
End Sub
Public Sub LkValToggle18()
    LkValToggle 18
End Sub
Public Sub LkValToggle19()
    LkValToggle 19
End Sub
Public Sub LkValToggle20()
    LkValToggle 20
End Sub
Public Sub LkValToggle21()
    LkValToggle 21
End Sub
Public Sub LkValToggle22()
    LkValToggle 22
End Sub
Public Sub LkValToggle23()
    LkValToggle 23
End Sub
Public Sub LkValToggle24()
    LkValToggle 24
End Sub
Public Sub LkValToggle25()
    LkValToggle 25
End Sub
Public Sub LkValToggle26()
    LkValToggle 26
End Sub
Public Sub LkValToggle27()
    LkValToggle 27
End Sub
Public Sub LkValToggle28()
    LkValToggle 28
End Sub
Public Sub LkValToggle29()
    LkValToggle 29
End Sub
Public Sub LkValToggle30()
    LkValToggle 30
End Sub
Public Sub LkValToggle31()
    LkValToggle 31
End Sub
Public Sub LkValToggle32()
    LkValToggle 32
End Sub
Public Sub LkValToggle33()
    LkValToggle 33
End Sub
Public Sub LkValToggle34()
    LkValToggle 34
End Sub
Public Sub LkValToggle35()
    LkValToggle 35
End Sub
Public Sub LkValToggle36()
    LkValToggle 36
End Sub
Public Sub LkValToggle37()
    LkValToggle 37
End Sub
Public Sub LkValToggle38()
    LkValToggle 38
End Sub
Public Sub LkValToggle39()
    LkValToggle 39
End Sub
Public Sub LkValToggle40()
    LkValToggle 40
End Sub
Public Sub LkValToggle41()
    LkValToggle 41
End Sub
Public Sub LkValToggle42()
    LkValToggle 42
End Sub
Public Sub LkValToggle43()
    LkValToggle 43
End Sub
Public Sub LkValToggle44()
    LkValToggle 44
End Sub
Public Sub LkValToggle45()
    LkValToggle 45
End Sub
Public Sub LkValToggle46()
    LkValToggle 46
End Sub
Public Sub LkValToggle47()
    LkValToggle 47
End Sub
Public Sub LkValToggle48()
    LkValToggle 48
End Sub
Public Sub LkValToggle49()
    LkValToggle 49
End Sub
Public Sub LkValToggle50()
    LkValToggle 50
End Sub
Public Sub LkValToggle51()
    LkValToggle 51
End Sub
Public Sub LkValToggle52()
    LkValToggle 52
End Sub
Public Sub LkValToggle53()
    LkValToggle 53
End Sub
Public Sub LkValToggle54()
    LkValToggle 54
End Sub
Public Sub LkValToggle55()
    LkValToggle 55
End Sub
Public Sub LkValToggle56()
    LkValToggle 56
End Sub
Public Sub LkValToggle57()
    LkValToggle 57
End Sub
Public Sub LkValToggle58()
    LkValToggle 58
End Sub
Public Sub LkValToggle59()
    LkValToggle 59
End Sub
Public Sub LkValToggle60()
    LkValToggle 60
End Sub

'============================================================
'  column headers to row data  (Ctrl+Alt+B)
'    columns spread sideways, like Jan Feb Mar, are stood up into one long column.
'    (the same job as Power Query's "Unpivot Columns")
'
'      ID       Name    Jan  Feb  Mar
'      00123  J. Smith  10   20   30
'            ->
'      ID       Name    Item  Value
'      00123  J. Smith  Jan    10
'      00123  J. Smith  Feb    20
'      00123  J. Smith  Mar    30
'
'    the popup shows one single list.
'      a + column moves down into rows, an o column stays where it is.
'============================================================
Public Sub Unpivot()
    Dim ws As Worksheet
    Dim data As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Column Headers to Row Data"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set data = PickColumnRange(ws)
    If data Is Nothing Then Exit Sub

    If StrComp(mUpWs, ws.Name, vbTextCompare) <> 0 Or _
       StrComp(mUpAddr, data.Address, vbTextCompare) <> 0 Then
        ReDim mUpVal(1 To data.Columns.Count)
        mUpWs = ws.Name
        mUpAddr = data.Address
        mUpBlank = True                       ' default : blanks are included as they are
        mUpDown = False                       ' default : pick the columns to move
    End If

    If Len(mUpName1) = 0 Then mUpName1 = "Item"
    If Len(mUpName2) = 0 Then mUpName2 = "Value"

    On Error Resume Next
    If UBound(mUpVal) <> data.Columns.Count Then
        ReDim mUpVal(1 To data.Columns.Count)
    End If
    Err.Clear
    On Error GoTo 0

    UnpivotMenu data
End Sub

'--- the column-picking popup (one single list) ---
Private Sub UnpivotMenu(ByVal data As Range)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim i As Long
    Dim colCount As Long, unpivotCount2 As Long
    Dim name2 As String
    Dim sourceRows As Double
    Dim approx As String

    On Error Resume Next
    Application.CommandBars(BAR_UNPIVOT).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=BAR_UNPIVOT, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    colCount = data.Columns.Count
    For i = 1 To colCount
        If mUpVal(i) Then unpivotCount2 = unpivotCount2 + 1
    Next i
    sourceRows = CDbl(data.Rows.Count - 1)

    ' the number drops when rows are hidden by a filter or blanks are skipped
    approx = ""
    On Error Resume Next
    If data.Worksheet.FilterMode Then approx = "about "
    Err.Clear
    On Error GoTo Fail
    If Not mUpBlank Then approx = "about "

    AddMenuTitle cb, "Column Headers to Row Data        (" & data.Address(False, False) & ")"
    AddMenuTitle cb, "●  marked columns move down into rows.   ○  stays as is."

    For i = 1 To colCount
        name2 = ColumnLabel(data, i)
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mUpVal(i), "     ●  ", "     ○  ") & name2
        btn.OnAction = MacroPrefix() & "UnpivotToggle" & Format$(i, "00")
        btn.Style = msoButtonCaption
        If i = 1 Then btn.BeginGroup = True
    Next i

    '--- run ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Create on a new sheet   (" & Format$(sourceRows, "#,##0") & " rows  ->  " & _
                  approx & Format$(sourceRows * CDbl(unpivotCount2), "#,##0") & " rows)"
    btn.OnAction = MacroPrefix() & "RunUnpivot"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True
    btn.Enabled = (unpivotCount2 > 0 And unpivotCount2 < colCount)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Blank cells : " & IIf(mUpBlank, "included  ->  skipped", "skip  ->  include")
    btn.OnAction = MacroPrefix() & "ToggleUnpivotBlanks"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     New column names : " & mUpName1 & " / " & mUpName2 & "        (click to change)"
    btn.OnAction = MacroPrefix() & "RenameUnpivotColumns"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Picking mode : " & _
                  IIf(mUpDown, "pick columns to keep  ->  pick columns to move", _
                               "pick columns to move  ->  pick columns to keep")
    btn.OnAction = MacroPrefix() & "ToggleUnpivotPickMode"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Clear all        (columns to move : " & unpivotCount2 & ")"
    btn.OnAction = MacroPrefix() & "UnpivotClearAll"
    btn.Style = msoButtonCaption

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "Could not show the menu." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Column Headers to Row Data"
End Sub

Public Sub ToggleUnpivotBlanks()
    mUpBlank = Not mUpBlank
    Unpivot
End Sub

'--- switch the picking mode : every + and o swaps over ---
Public Sub ToggleUnpivotPickMode()
    Dim i As Long
    On Error Resume Next
    For i = LBound(mUpVal) To UBound(mUpVal)
        mUpVal(i) = Not mUpVal(i)
    Next i
    mUpDown = Not mUpDown
    Err.Clear
    On Error GoTo 0
    Unpivot
End Sub

Public Sub UnpivotClearAll()
    Dim i As Long
    On Error Resume Next
    For i = LBound(mUpVal) To UBound(mUpVal)
        mUpVal(i) = False
    Next i
    mUpDown = False
    Err.Clear
    On Error GoTo 0
    Unpivot
End Sub

Public Sub RenameUnpivotColumns()
    Dim s As String

    s = InputBox("Type the names of the two new columns, separated by a comma." & vbCrLf & vbCrLf & _
                 "First = name of the column holding the headers" & vbCrLf & _
                 "Second = name of the column holding the values", _
                 "New column names", mUpName1 & "," & mUpName2)
    If Len(Trim$(s)) = 0 Then
        Unpivot
        Exit Sub
    End If

    If InStr(1, s, ",") > 0 Then
        mUpName1 = Trim$(Split(s, ",")(0))
        mUpName2 = Trim$(Split(s, ",")(1))
    Else
        mUpName1 = Trim$(s)
    End If
    If Len(mUpName1) = 0 Then mUpName1 = "Item"
    If Len(mUpName2) = 0 Then mUpName2 = "Value"

    Unpivot
End Sub

Private Sub UnpivotToggle(ByVal no As Long)
    On Error Resume Next
    mUpVal(no) = Not mUpVal(no)
    Err.Clear
    On Error GoTo 0
    Unpivot
End Sub

'------------------------------------------------------------
' the conversion itself - onto a new sheet
'------------------------------------------------------------
Public Sub RunUnpivot()
    Dim ws As Worksheet, newSheet As Worksheet
    Dim data As Range
    Dim visibleRange As Range, area As Range
    Dim visible() As Boolean
    Dim keepCols() As Long, unpivotCols() As Long
    Dim keepCount As Long, unpivotCount As Long
    Dim grid() As Variant, oneCol As Variant
    Dim header() As String
    Dim buf() As Variant, tailBuf() As Variant
    Dim rowCnt As Long, visibleCount As Long
    Dim r As Long, c As Long, i As Long, j As Long, startIdx As Long
    Dim bufSize As Long, bufRow As Long, writtenRows As Long
    Dim resultRows As Double, rowsWithBlanks As Double
    Dim name2 As String, numFmt As String
    Dim blankCount As Double
    Dim v As Variant
    Dim keepIt As Boolean
    Dim colTotal As Long
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldCalc As Long
    Dim locked As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Len(mUpAddr) = 0 Then
        MsgBox "Press Ctrl+Alt+B first to pick the columns.", _
               vbInformation, "Column Headers to Row Data"
        Exit Sub
    End If
    Set ws = ActiveSheet
    Set data = ws.Range(mUpAddr)
    rowCnt = data.Rows.Count

    '--- + marks a column to move, everything else is a key column ---
    ReDim keepCols(1 To data.Columns.Count)
    ReDim unpivotCols(1 To data.Columns.Count)
    For c = 1 To data.Columns.Count
        If mUpVal(c) Then
            unpivotCount = unpivotCount + 1
            unpivotCols(unpivotCount) = c
        Else
            keepCount = keepCount + 1
            keepCols(keepCount) = c
        End If
    Next c

    If unpivotCount < 1 Then
        MsgBox "Please pick at least one column to move into rows." & vbCrLf & _
               "Columns marked ● move down into rows.", _
               vbExclamation, "Column Headers to Row Data"
        Exit Sub
    End If
    If keepCount < 1 Then
        MsgBox "There are no columns left to keep as is." & vbCrLf & _
               "Please unmark at least one column so it stays as ○.", _
               vbExclamation, "Column Headers to Row Data"
        Exit Sub
    End If

    '--- visible rows only (filters taken into account) ---
    ReDim visible(1 To rowCnt)
    On Error Resume Next
    Set visibleRange = data.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo Fail
    If visibleRange Is Nothing Then
        For r = 2 To rowCnt
            visible(r) = True
            visibleCount = visibleCount + 1
        Next r
    Else
        For Each area In visibleRange.Areas
            startIdx = area.Row - data.Row + 1
            For r = startIdx To startIdx + area.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not visible(r) Then
                        visible(r) = True
                        visibleCount = visibleCount + 1
                    End If
                End If
            Next r
        Next area
    End If

    If visibleCount < 1 Then
        MsgBox "There are no visible data rows.", vbExclamation, "Column Headers to Row Data"
        Exit Sub
    End If

    '--- read only the columns needed into an array (saves memory) ---
    Application.StatusBar = "Reading the data..."
    ReDim grid(1 To rowCnt, 1 To keepCount + unpivotCount)
    For i = 1 To keepCount
        oneCol = data.Columns(keepCols(i)).Value
        For r = 1 To rowCnt
            grid(r, i) = oneCol(r, 1)
        Next r
    Next i
    ReDim header(1 To unpivotCount)
    For i = 1 To unpivotCount
        oneCol = data.Columns(unpivotCols(i)).Value
        For r = 1 To rowCnt
            grid(r, keepCount + i) = oneCol(r, 1)
        Next r
        name2 = Trim$(CStr(data.Cells(1, unpivotCols(i)).Value))
        If Len(name2) = 0 Then name2 = "Column " & ColumnLetter(data.Column + unpivotCols(i) - 1)
        header(i) = name2
    Next i
    oneCol = Empty

    '--- count the rows that will be created first ---
    Application.StatusBar = "Counting how many rows this will make..."
    rowsWithBlanks = CDbl(visibleCount) * CDbl(unpivotCount)
    For r = 2 To rowCnt
        If visible(r) Then
            For i = 1 To unpivotCount
                v = grid(r, keepCount + i)
                If IsEmpty(v) Then
                    blankCount = blankCount + 1
                ElseIf VarType(v) = vbString Then
                    If Len(v) = 0 Then blankCount = blankCount + 1
                End If
            Next i
        End If
    Next r
    resultRows = rowsWithBlanks - blankCount
    If mUpBlank Then resultRows = rowsWithBlanks

    Application.StatusBar = False

    If resultRows < 1 Then
        MsgBox "There are no rows to create." & vbCrLf & _
               "Every column to move is empty.", vbExclamation, "Column Headers to Row Data"
        Exit Sub
    End If

    '--- stop when it would not fit on one sheet ---
    If resultRows + 3 > CDbl(ws.Rows.Count) Then
        MsgBox "This would create " & Format$(resultRows, "#,##0") & " rows" & vbCrLf & _
               "one sheet (" & Format$(ws.Rows.Count, "#,##0") & " rows)." & vbCrLf & vbCrLf & _
               "Pick fewer columns to move, or filter the source and" & vbCrLf & _
               "leave only the visible rows, then run again.", _
               vbExclamation, "Column Headers to Row Data"
        Exit Sub
    End If

    '--- confirm before running ---
    If MsgBox("On a new sheet, " & Format$(resultRows, "#,##0") & " rows will be created." & vbCrLf & vbCrLf & _
              "Source " & Format$(visibleCount, "#,##0") & " rows" & _
              IIf(visibleCount < rowCnt - 1, "  (visible rows only)", "") & vbCrLf & _
              "Columns kept as is : " & keepCount & "" & vbCrLf & _
              "Columns to move into rows : " & unpivotCount & "" & vbCrLf & vbCrLf & _
              "Including blanks : " & Format$(rowsWithBlanks, "#,##0") & " rows" & vbCrLf & _
              "Skipping blanks : " & Format$(rowsWithBlanks - blankCount, "#,##0") & " rows" & vbCrLf & _
              "Current setting : " & IIf(mUpBlank, "include blanks", "skip blanks") & vbCrLf & vbCrLf & _
              IIf(resultRows > 100000, "That is a lot of rows. It is safer to save the file first." & vbCrLf & vbCrLf, "") & _
              "Do you want to continue?", _
              vbOKCancel + vbQuestion, "Column Headers to Row Data") <> vbOK Then
        Exit Sub
    End If

    '--- lock ---
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    locked = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    colTotal = keepCount + 2

    '--- the new sheet ---
    name2 = NewSheetName(ws.Parent, "Unpivot_" & ws.Name)
    Set newSheet = ws.Parent.Worksheets.Add(Before:=ws)
    On Error Resume Next
    newSheet.Name = name2
    Err.Clear
    On Error GoTo Fail

    With newSheet.Range("A1")
        .Value = "Column Headers to Row Data  |  source : " & ws.Name & "!" & data.Address(False, False) & _
                 "  |  columns moved : " & unpivotCount & "" & _
                 IIf(mUpBlank, "  |  blanks included", "  |  blanks skipped")
        .Font.Size = 9
        .Font.Color = RGB(110, 110, 115)
    End With
    newSheet.Columns("A").ColumnWidth = 12

    '--- move the number formats across before the values ---
    For i = 1 To keepCount
        newSheet.Cells(4, i).Resize(CLng(resultRows), 1).NumberFormat = _
            data.Cells(2, keepCols(i)).NumberFormat
    Next i
    numFmt = data.Cells(2, unpivotCols(1)).NumberFormat
    For i = 2 To unpivotCount
        If data.Cells(2, unpivotCols(i)).NumberFormat <> numFmt Then
            numFmt = "General"
            Exit For
        End If
    Next i
    newSheet.Cells(4, keepCount + 1).Resize(CLng(resultRows), 1).NumberFormat = "@"
    newSheet.Cells(4, colTotal).Resize(CLng(resultRows), 1).NumberFormat = numFmt

    '--- the headers ---
    ReDim buf(1 To 1, 1 To colTotal)
    For i = 1 To keepCount
        buf(1, i) = grid(1, i)
    Next i
    buf(1, keepCount + 1) = mUpName1
    buf(1, colTotal) = mUpName2
    newSheet.Range("A3").Resize(1, colTotal).Value = buf
    With newSheet.Range("A3").Resize(1, colTotal)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    '--- write the body in chunks ---
    bufSize = 50000
    If bufSize > CLng(resultRows) Then bufSize = CLng(resultRows)
    ReDim buf(1 To bufSize, 1 To colTotal)
    bufRow = 0
    writtenRows = 0

    For r = 2 To rowCnt
      If visible(r) Then
        For i = 1 To unpivotCount
            v = grid(r, keepCount + i)
            keepIt = True
            If Not mUpBlank Then
                If IsEmpty(v) Then
                    keepIt = False
                ElseIf VarType(v) = vbString Then
                    If Len(v) = 0 Then keepIt = False
                End If
            End If

            If keepIt Then
                bufRow = bufRow + 1
                For c = 1 To keepCount
                    buf(bufRow, c) = grid(r, c)
                Next c
                buf(bufRow, keepCount + 1) = header(i)
                buf(bufRow, colTotal) = v

                If bufRow = bufSize Then
                    newSheet.Cells(4 + writtenRows, 1).Resize(bufSize, colTotal).Value = buf
                    writtenRows = writtenRows + bufSize
                    bufRow = 0
                    Application.StatusBar = "Creating...  " & _
                        Format$(writtenRows, "#,##0") & " / " & Format$(resultRows, "#,##0") & " rows"
                End If
            End If
        Next i
      End If
    Next r

    '--- what is left ---
    If bufRow > 0 Then
        ReDim tailBuf(1 To bufRow, 1 To colTotal)
        For j = 1 To bufRow
            For c = 1 To colTotal
                tailBuf(j, c) = buf(j, c)
            Next c
        Next j
        newSheet.Cells(4 + writtenRows, 1).Resize(bufRow, colTotal).Value = tailBuf
        writtenRows = writtenRows + bufRow
    End If

    newSheet.Range("A3").Resize(1, colTotal).Columns.AutoFit
    newSheet.Activate
    newSheet.Range("A4").Select

    '--- unlock ---
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    locked = False
    Application.StatusBar = False

    MsgBox Format$(writtenRows, "#,##0") & " rows created." & vbCrLf & vbCrLf & _
           "Sheet : " & name2 & vbCrLf & _
           "Columns : kept " & keepCount & " + " & mUpName1 & " + " & mUpName2 & vbCrLf & _
           "Source : " & ws.Name & "!" & data.Address(False, False) & _
           IIf(visibleCount < rowCnt - 1, "  (visible " & Format$(visibleCount, "#,##0") & " rows only)", ""), _
           vbInformation, "Column Headers to Row Data"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    If locked Then
        Application.Calculation = oldCalc
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    Application.StatusBar = False
    MsgBox "Could not create it." & vbCrLf & _
           "Error " & errNo & ": " & errDesc & vbCrLf & vbCrLf & _
           "If there are too many rows, pick fewer columns or" & vbCrLf & _
           "split the source and run it in parts.", vbExclamation, "Column Headers to Row Data"
End Sub

'--- toggle wrappers for the columns to move ---
Public Sub UnpivotToggle01()
    UnpivotToggle 1
End Sub
Public Sub UnpivotToggle02()
    UnpivotToggle 2
End Sub
Public Sub UnpivotToggle03()
    UnpivotToggle 3
End Sub
Public Sub UnpivotToggle04()
    UnpivotToggle 4
End Sub
Public Sub UnpivotToggle05()
    UnpivotToggle 5
End Sub
Public Sub UnpivotToggle06()
    UnpivotToggle 6
End Sub
Public Sub UnpivotToggle07()
    UnpivotToggle 7
End Sub
Public Sub UnpivotToggle08()
    UnpivotToggle 8
End Sub
Public Sub UnpivotToggle09()
    UnpivotToggle 9
End Sub
Public Sub UnpivotToggle10()
    UnpivotToggle 10
End Sub
Public Sub UnpivotToggle11()
    UnpivotToggle 11
End Sub
Public Sub UnpivotToggle12()
    UnpivotToggle 12
End Sub
Public Sub UnpivotToggle13()
    UnpivotToggle 13
End Sub
Public Sub UnpivotToggle14()
    UnpivotToggle 14
End Sub
Public Sub UnpivotToggle15()
    UnpivotToggle 15
End Sub
Public Sub UnpivotToggle16()
    UnpivotToggle 16
End Sub
Public Sub UnpivotToggle17()
    UnpivotToggle 17
End Sub
Public Sub UnpivotToggle18()
    UnpivotToggle 18
End Sub
Public Sub UnpivotToggle19()
    UnpivotToggle 19
End Sub
Public Sub UnpivotToggle20()
    UnpivotToggle 20
End Sub
Public Sub UnpivotToggle21()
    UnpivotToggle 21
End Sub
Public Sub UnpivotToggle22()
    UnpivotToggle 22
End Sub
Public Sub UnpivotToggle23()
    UnpivotToggle 23
End Sub
Public Sub UnpivotToggle24()
    UnpivotToggle 24
End Sub
Public Sub UnpivotToggle25()
    UnpivotToggle 25
End Sub
Public Sub UnpivotToggle26()
    UnpivotToggle 26
End Sub
Public Sub UnpivotToggle27()
    UnpivotToggle 27
End Sub
Public Sub UnpivotToggle28()
    UnpivotToggle 28
End Sub
Public Sub UnpivotToggle29()
    UnpivotToggle 29
End Sub
Public Sub UnpivotToggle30()
    UnpivotToggle 30
End Sub
Public Sub UnpivotToggle31()
    UnpivotToggle 31
End Sub
Public Sub UnpivotToggle32()
    UnpivotToggle 32
End Sub
Public Sub UnpivotToggle33()
    UnpivotToggle 33
End Sub
Public Sub UnpivotToggle34()
    UnpivotToggle 34
End Sub
Public Sub UnpivotToggle35()
    UnpivotToggle 35
End Sub
Public Sub UnpivotToggle36()
    UnpivotToggle 36
End Sub
Public Sub UnpivotToggle37()
    UnpivotToggle 37
End Sub
Public Sub UnpivotToggle38()
    UnpivotToggle 38
End Sub
Public Sub UnpivotToggle39()
    UnpivotToggle 39
End Sub
Public Sub UnpivotToggle40()
    UnpivotToggle 40
End Sub
Public Sub UnpivotToggle41()
    UnpivotToggle 41
End Sub
Public Sub UnpivotToggle42()
    UnpivotToggle 42
End Sub
Public Sub UnpivotToggle43()
    UnpivotToggle 43
End Sub
Public Sub UnpivotToggle44()
    UnpivotToggle 44
End Sub
Public Sub UnpivotToggle45()
    UnpivotToggle 45
End Sub
Public Sub UnpivotToggle46()
    UnpivotToggle 46
End Sub
Public Sub UnpivotToggle47()
    UnpivotToggle 47
End Sub
Public Sub UnpivotToggle48()
    UnpivotToggle 48
End Sub
Public Sub UnpivotToggle49()
    UnpivotToggle 49
End Sub
Public Sub UnpivotToggle50()
    UnpivotToggle 50
End Sub
Public Sub UnpivotToggle51()
    UnpivotToggle 51
End Sub
Public Sub UnpivotToggle52()
    UnpivotToggle 52
End Sub
Public Sub UnpivotToggle53()
    UnpivotToggle 53
End Sub
Public Sub UnpivotToggle54()
    UnpivotToggle 54
End Sub
Public Sub UnpivotToggle55()
    UnpivotToggle 55
End Sub
Public Sub UnpivotToggle56()
    UnpivotToggle 56
End Sub
Public Sub UnpivotToggle57()
    UnpivotToggle 57
End Sub
Public Sub UnpivotToggle58()
    UnpivotToggle 58
End Sub
Public Sub UnpivotToggle59()
    UnpivotToggle 59
End Sub
Public Sub UnpivotToggle60()
    UnpivotToggle 60
End Sub

'--- what kind of value it starts from ---
Private Function SeriesKind(ByVal s As String) As String
    If IsNumeric(s) Then
        SeriesKind = "number"
    ElseIf IsDate(s) Then
        SeriesKind = "date"
    ElseIf WeekdayNumber(s, False) > 0 Then
        SeriesKind = "weekday"
    ElseIf WeekdayNumber(s, True) > 0 Then
        SeriesKind = "weekdayLong"
    ElseIf MonthNumber(s, False) > 0 Then
        SeriesKind = "month"
    ElseIf MonthNumber(s, True) > 0 Then
        SeriesKind = "monthLong"
    ElseIf HeadNumber(s, "Q") > 0 Then
        SeriesKind = "quarter"
    ElseIf HeadNumber(s, "Week ") > 0 Then
        SeriesKind = "week"
    ElseIf HeadNumber(s, "Day ") > 0 Then
        SeriesKind = "day"
    ElseIf LeadingNumber(s) > 0 Then
        SeriesKind = "leadingNum"
    Else
        SeriesKind = "copy"
    End If
End Function

'--- the nth value ---
Private Function SeriesValue(ByVal seedVal As String, ByVal kind As String, _
                          ByVal n As Long, ByVal stepVal As Double) As Variant
    Dim i As Long
    Dim suffix As String

    Select Case kind
        Case "number"
            SeriesValue = CDbl(seedVal) + n * stepVal

        Case "date"
            SeriesValue = CDate(seedVal) + n * stepVal

        Case "weekday"
            SeriesValue = WeekdayName2(WeekdayNumber(seedVal, False) + n, False)

        Case "weekdayLong"
            SeriesValue = WeekdayName2(WeekdayNumber(seedVal, True) + n, True)

        Case "month"
            SeriesValue = MonthName2(MonthNumber(seedVal, False) + n, False)

        Case "monthLong"
            SeriesValue = MonthName2(MonthNumber(seedVal, True) + n, True)

        Case "quarter"
            i = HeadNumber(seedVal, "Q") + n
            i = ((i - 1) Mod 4 + 4) Mod 4 + 1
            SeriesValue = "Q" & i

        Case "week"
            SeriesValue = "Week " & (HeadNumber(seedVal, "Week ") + n)

        Case "day"
            SeriesValue = "Day " & (HeadNumber(seedVal, "Day ") + n)

        Case "leadingNum"
            suffix = Mid$(seedVal, Len(CStr(LeadingNumber(seedVal))) + 1)
            SeriesValue = (LeadingNumber(seedVal) + n) & suffix

        Case Else
            SeriesValue = seedVal
    End Select
End Function

'--- a weekday name -> 1 (Sun) ~ 7 (Sat), otherwise 0 ---
Private Function WeekdayNumber(ByVal s As String, ByVal longForm As Boolean) As Long
    Dim names As Variant
    Dim i As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If longForm Then
        names = Array("Sunday", "Monday", "Tuesday", "Wednesday", _
                      "Thursday", "Friday", "Saturday")
    Else
        names = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    End If
    For i = 0 To 6
        If StrComp(s, names(i), vbTextCompare) = 0 Then
            WeekdayNumber = i + 1
            Exit Function
        End If
    Next i
End Function

'--- Jan Feb ... / January February ... ---
Private Function MonthNumber(ByVal s As String, ByVal longForm As Boolean) As Long
    Dim names As Variant
    Dim i As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If longForm Then
        names = Array("January", "February", "March", "April", "May", "June", _
                      "July", "August", "September", "October", "November", "December")
    Else
        names = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    End If
    For i = 0 To 11
        If StrComp(s, names(i), vbTextCompare) = 0 Then
            MonthNumber = i + 1
            Exit Function
        End If
    Next i
End Function

'--- 이름으로 돌려준다 ---
Private Function WeekdayName2(ByVal n As Long, ByVal longForm As Boolean) As String
    Dim names As Variant
    If longForm Then
        names = Array("Sunday", "Monday", "Tuesday", "Wednesday", _
                      "Thursday", "Friday", "Saturday")
    Else
        names = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    End If
    WeekdayName2 = names(((n - 1) Mod 7 + 7) Mod 7)
End Function

Private Function MonthName2(ByVal n As Long, ByVal longForm As Boolean) As String
    Dim names As Variant
    If longForm Then
        names = Array("January", "February", "March", "April", "May", "June", _
                      "July", "August", "September", "October", "November", "December")
    Else
        names = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                      "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    End If
    MonthName2 = names(((n - 1) Mod 12 + 12) Mod 12)
End Function

'--- Q1 · W3 처럼 머리 글자 + 숫자 ---
Private Function HeadNumber(ByVal s As String, ByVal head As String) As Long
    Dim tailPart As String
    s = Trim$(s)
    If Len(s) <= Len(head) Then Exit Function
    If StrComp(Left$(s, Len(head)), head, vbTextCompare) <> 0 Then Exit Function
    tailPart = Mid$(s, Len(head) + 1)
    If Not IsNumeric(tailPart) Then Exit Function
    If CDbl(tailPart) <> Int(CDbl(tailPart)) Then Exit Function
    HeadNumber = CLng(tailPart)
End Function

'--- "Mar" -> 3,  "Q2" -> 2.  otherwise 0 ---
Private Function TrailingNumber(ByVal s As String, ByVal suffix As String) As Long
    Dim headPart As String
    s = Trim$(s)
    If Len(s) <= Len(suffix) Then Exit Function
    If Right$(s, Len(suffix)) <> suffix Then Exit Function
    headPart = Left$(s, Len(s) - Len(suffix))
    If Not IsNumeric(headPart) Then Exit Function
    If CDbl(headPart) <> Int(CDbl(headPart)) Then Exit Function
    TrailingNumber = CLng(headPart)
End Function

'--- a leading number, as in "Round 1" or "Team 2". otherwise 0 ---
Private Function LeadingNumber(ByVal s As String) As Long
    Dim i As Long
    Dim numCount As String
    s = Trim$(s)
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then
            numCount = numCount & Mid$(s, i, 1)
        Else
            Exit For
        End If
    Next i
    If Len(numCount) = 0 Then Exit Function
    If Len(numCount) = Len(s) Then Exit Function
    LeadingNumber = CLng(numCount)
End Function

'============================================================
'  center / default (Ctrl+Alt+Z)
'============================================================
Public Sub ToggleCenterAlign()
    Dim rng As Range
    Dim nowVal As Variant
    Dim toCenter As Boolean
    Dim errNo As Long, errDesc As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell first.", vbExclamation, "Center Align"
        Exit Sub
    End If
    Set rng = Selection

    On Error GoTo Fail

    nowVal = rng.HorizontalAlignment
    toCenter = True
    If Not IsNull(nowVal) Then
        If nowVal = xlCenter Then toCenter = False
    End If

    SaveAlignment rng

    If toCenter Then
        rng.HorizontalAlignment = xlCenter
    Else
        rng.HorizontalAlignment = xlGeneral
    End If

    If mAlignOn Then
        On Error Resume Next
        Application.OnUndo "Undo Alignment", MacroPrefix() & "UndoAlignment"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = IIf(toCenter, "Centered.", "Restored the default alignment.") & _
                            "  (" & rng.Address(False, False) & ")" & _
                            IIf(mAlignOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "Could not change the alignment." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Center Align"
End Sub

'============================================================
'  center across selection (Ctrl+Alt+S)
'    it looks spanned without the cells being merged.
'    unlike merging, sorting, filtering and pivots keep working.
'============================================================
Public Sub CenterAcrossSelection()
    Dim rng As Range
    Dim nowVal As Variant
    Dim toAcross As Boolean
    Dim errNo As Long, errDesc As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell first.", vbExclamation, "Center Across Selection"
        Exit Sub
    End If
    Set rng = Selection

    If rng.Columns.Count < 2 Then
        MsgBox "Please select two or more cells across." & vbCrLf & vbCrLf & _
               "The text in the leftmost cell is shown centered across the selection." & vbCrLf & _
               "Cells are not merged, so sorting, filtering and pivots still work.", _
               vbExclamation, "Center Across Selection"
        Exit Sub
    End If

    On Error GoTo Fail

    nowVal = rng.HorizontalAlignment
    toAcross = True
    If Not IsNull(nowVal) Then
        If nowVal = xlCenterAcrossSelection Then toAcross = False
    End If

    SaveAlignment rng

    If toAcross Then
        rng.HorizontalAlignment = xlCenterAcrossSelection
    Else
        rng.HorizontalAlignment = xlGeneral
    End If

    If mAlignOn Then
        On Error Resume Next
        Application.OnUndo "Undo Alignment", MacroPrefix() & "UndoAlignment"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = IIf(toAcross, "Centered across the selection. (cells are not merged)", _
                                          "Restored the default alignment.") & _
                            IIf(mAlignOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "Could not change it." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Center Across Selection"
End Sub

'--- kept so the alignment can be undone ---
Private Sub SaveAlignment(ByVal rg As Range)
    Dim r As Long, c As Long
    Dim v() As Variant

    mAlignOn = False
    If rg Is Nothing Then Exit Sub
    If rg.Areas.Count > 1 Then Exit Sub
    If rg.Cells.CountLarge > 30000 Then Exit Sub

    On Error GoTo Fail
    ReDim v(1 To rg.Rows.Count, 1 To rg.Columns.Count)
    For r = 1 To rg.Rows.Count
        For c = 1 To rg.Columns.Count
            v(r, c) = rg.Cells(r, c).HorizontalAlignment
        Next c
    Next r

    mAlignWb = rg.Worksheet.Parent.Name
    mAlignWs = rg.Worksheet.Name
    mAlignAddr = rg.Address
    mAlignVal = v
    mAlignOn = True
    Exit Sub
Fail:
    mAlignOn = False
End Sub

Public Sub UndoAlignment()
    Dim rng As Range
    Dim r As Long, c As Long
    Dim oldScreen As Boolean

    If Not mAlignOn Or Len(mAlignAddr) = 0 Then
        MsgBox "There is no alignment to undo.", vbInformation, "Alignment"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False
    Set rng = Application.Workbooks(mAlignWb).Worksheets(mAlignWs).Range(mAlignAddr)
    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            rng.Cells(r, c).HorizontalAlignment = mAlignVal(r, c)
        Next c
    Next r
    mAlignOn = False
    mAlignAddr = ""
    Err.Clear
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    Application.StatusBar = "The alignment was restored."
    ScheduleStatusClear
End Sub

'============================================================
'  fill blanks from above (Ctrl+Shift+4)
'    blanks in an unmerged table or in data pulled from a system
'    are flowed down from the value just above.
'============================================================
Public Sub FillBlanksFromAbove()
    Dim ws As Worksheet
    Dim rng As Range
    Dim blk As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim filledCount As Long
    Dim oldScreen As Boolean, oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the range whose blanks should be filled, then run again.", vbExclamation, "Fill Blanks From Above"
        Exit Sub
    End If

    If Selection.Cells.CountLarge >= 2 And Selection.Areas.Count = 1 Then
        Set rng = Selection
    Else
        Set blk = SkipNonHeaderRows(ActiveCell.CurrentRegion)
        If blk Is Nothing Then
            MsgBox "Could not find a table.", vbExclamation, "Fill Blanks From Above"
            Exit Sub
        End If
        Set rng = ws.Range(ws.Cells(blk.Row + 1, ActiveCell.Column), _
                           ws.Cells(blk.Row + blk.Rows.Count - 1, ActiveCell.Column))
    End If

    If rng Is Nothing Then
        MsgBox "There is no data to fill.", vbExclamation, "Fill Blanks From Above"
        Exit Sub
    End If
    If rng.Rows.Count < 2 Then
        MsgBox "Please select two or more rows.", vbExclamation, "Fill Blanks From Above"
        Exit Sub
    End If
    If rng.Cells.CountLarge > 200000 Then
        MsgBox "The range is too large. Narrow it and run again.", vbExclamation, "Fill Blanks From Above"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    SaveCriteriaBlock rng
    arr = rng.Value

    For c = 1 To rng.Columns.Count
        For r = 2 To rng.Rows.Count
            If Len(CStr(arr(r, c))) = 0 Then
                If Len(CStr(arr(r - 1, c))) > 0 Then
                    arr(r, c) = arr(r - 1, c)
                    filledCount = filledCount + 1
                End If
            End If
        Next r
    Next c

    If filledCount > 0 Then rng.Value = arr

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    If filledCount = 0 Then
        MsgBox "There are no blank cells to fill." & vbCrLf & _
               "Range : " & rng.Address(False, False), vbInformation, "Fill Blanks From Above"
        Exit Sub
    End If

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "Undo Fill Blanks", MacroPrefix() & "UndoCriteriaBlock"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "blank " & Format$(filledCount, "#,##0") & " blanks filled from the cell above. (" & _
                            rng.Address(False, False) & ")" & _
                            IIf(mCondOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
    MsgBox "Could not fill." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Fill Blanks From Above"
End Sub

'============================================================
'  remove the rows above the header (Ctrl+Alt+E)
'    clears the clutter above the header before exporting to text.
'============================================================
Public Sub RemoveRowsAboveHeader()
    Dim ws As Worksheet
    Dim blk As Range
    Dim headerRow As Long
    Dim removeCount As Long
    Dim preview As String
    Dim i As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    On Error GoTo Fail

    Set blk = SkipNonHeaderRows(ActiveCell.CurrentRegion)
    If blk Is Nothing Then
        MsgBox "Could not find a table.", vbExclamation, "Remove Rows Above Header"
        Exit Sub
    End If

    headerRow = blk.Row
    removeCount = headerRow - 1

    If removeCount < 1 Then
        MsgBox "The header is already in row 1." & vbCrLf & _
               "There are no rows to remove.", vbInformation, "Remove Rows Above Header"
        Exit Sub
    End If

    For i = 1 To removeCount
        If i > 5 Then
            preview = preview & "   ..." & vbCrLf
            Exit For
        End If
        preview = preview & "   " & i & "Rows : " & RowPreview(ws, i) & vbCrLf
    Next i

    If MsgBox("Above the header : " & removeCount & " rows will be removed." & vbCrLf & vbCrLf & _
              preview & vbCrLf & _
              "Row taken as header : " & headerRow & " rows  (" & RowPreview(ws, headerRow) & ")" & vbCrLf & vbCrLf & _
              "You can undo this with Ctrl+Z." & vbCrLf & vbCrLf & _
              "Do you want to continue?", vbOKCancel + vbQuestion, "Remove Rows Above Header") = vbCancel Then Exit Sub

    SaveRowsAboveHeader ws, removeCount

    ws.Range(ws.Rows(1), ws.Rows(removeCount)).Delete Shift:=xlUp

    If mDelOn Then
        On Error Resume Next
        Application.OnUndo "Undo Remove Rows Above Header", MacroPrefix() & "UndoRemoveRowsAboveHeader"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = removeCount & " rows removed. The header is now row 1." & _
                            IIf(mDelOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "Could not clear it." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Remove Rows Above Header"
End Sub

Private Function RowPreview(ByVal ws As Worksheet, ByVal rowIdx As Long) As String
    Dim c As Long
    Dim s As String
    Dim v As String
    On Error GoTo Done
    For c = 1 To 12
        v = Trim$(CStr(ws.Cells(rowIdx, c).Value))
        If Len(v) > 0 Then
            If Len(s) > 0 Then s = s & " | "
            s = s & Left$(v, 12)
        End If
        If Len(s) > 50 Then
            s = s & " ..."
            Exit For
        End If
    Next c
Done:
    If Len(s) = 0 Then s = "(blank row)"
    RowPreview = s
End Function

'============================================================
'  compare two columns (Ctrl+Alt+D)
'    whether each value in column A is in column B, and each in B is in A,
'    shown side by side as 1/0 and a count.
'============================================================
Public Sub CompareTwoColumns()
    Dim ws As Worksheet
    Dim data As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Compare Two Columns"
        Exit Sub
    End If
    Set ws = ActiveSheet

    '--- when two spots are picked with Ctrl, ask which job to do ---
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 2 Then
            mCmpDirA = Selection.Areas(1).Address
            mCmpDirB = Selection.Areas(2).Address
            mCmpDirWs = ws.Name
            TwoSpotMenu
            Exit Sub
        End If
    End If

    Set data = PickColumnRange(ws)
    If data Is Nothing Then Exit Sub

    If StrComp(mCmpWs, ws.Name, vbTextCompare) <> 0 Or _
       StrComp(mCmpAddr, data.Address, vbTextCompare) <> 0 Then
        mCmpWs = ws.Name
        mCmpAddr = data.Address
        mCmpA = 0
        mCmpB = 0
        mCmpFormula = True
    End If

    CompareColumnsMenu data
End Sub

'--- compare the two ranges picked with Ctrl straight away ---
Private Sub CompareDirect(ByVal rngA As Range, ByVal rngB As Range)
    Dim aName As String, bName As String
    Dim a As Range, b As Range

    Set a = rngA
    Set b = rngB
    TrimCompareRange a, b
    Set rngA = a
    Set rngB = b

    If rngA.Columns.Count > 1 Or rngB.Columns.Count > 1 Then
        MsgBox "Please select one column at a time." & vbCrLf & _
               "Hold Ctrl and click the two columns (or two ranges) you want to compare.", _
               vbExclamation, "Compare Two Columns"
        Exit Sub
    End If

    aName = RangeHeader(rngA)
    bName = RangeHeader(rngB)

    CompareCore rngA, rngB, aName, bName
End Sub

'--- find the header in the cells above the range ---
Private Function RangeHeader(ByVal rg As Range) As String
    Dim s As String
    On Error GoTo Done
    If rg.Row > 1 Then s = Trim$(CStr(rg.Cells(1, 1).Offset(-1, 0).Value))
Done:
    If Len(s) = 0 Then s = Split(rg.Cells(1, 1).Address(True, False), "$")(0) & "Column"
    RangeHeader = s
End Function

Private Sub CompareColumnsMenu(ByVal data As Range)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim i As Long

    On Error Resume Next
    Application.CommandBars(BAR_COMPARE_COLS).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=BAR_COMPARE_COLS, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    AddMenuTitle cb, "Compare Two Columns  -  pick a method below"

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "(1)  Column to check  (one)"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    For i = 1 To data.Columns.Count
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mCmpA = i, "     ■  ", "     □  ") & ColumnLabel(data, i)
        btn.OnAction = MacroPrefix() & "CmpAToggle" & Format$(i, "00")
        btn.Style = msoButtonCaption
    Next i

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "(2)  List to look in  (one)"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    For i = 1 To data.Columns.Count
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mCmpB = i, "     ●  ", "     ○  ") & ColumnLabel(data, i)
        btn.OnAction = MacroPrefix() & "CmpBToggle" & Format$(i, "00")
        btn.Style = msoButtonCaption
    Next i

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Result : " & IIf(mCmpFormula, "as formulas (COUNTIF, updates itself)", "as values (fast)")
    btn.OnAction = MacroPrefix() & "ToggleCompareFormula"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Check against the list   (4 result columns, both ways)"
    btn.OnAction = MacroPrefix() & "RunCompareTwoColumns"
    btn.Style = msoButtonCaption
    btn.Enabled = (mCmpA > 0 And mCmpB > 0 And mCmpA <> mCmpB)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Compare within each row   (mark and color the differing rows)"
    btn.OnAction = MacroPrefix() & "RunRowCompare"
    btn.Style = msoButtonCaption
    btn.Enabled = (mCmpA > 0 And mCmpB > 0 And mCmpA <> mCmpB)

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "Could not show the menu." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Compare Two Columns"
End Sub

Public Sub ToggleCompareFormula()
    mCmpFormula = Not mCmpFormula
    CompareTwoColumns
End Sub

Private Sub CmpAToggle(ByVal no As Long)
    mCmpA = no
    If mCmpB = no Then mCmpB = 0
    CompareTwoColumns
End Sub

Private Sub CmpBToggle(ByVal no As Long)
    mCmpB = no
    If mCmpA = no Then mCmpA = 0
    CompareTwoColumns
End Sub

Public Sub RunCompareTwoColumns()
    Dim ws As Worksheet
    Dim data As Range
    Dim rngA As Range, rngB As Range

    On Error GoTo Fail
    Set ws = ActiveSheet
    Set data = ws.Range(mCmpAddr)

    Set rngA = data.Columns(mCmpA).Offset(1, 0).Resize(data.Rows.Count - 1, 1)
    Set rngB = data.Columns(mCmpB).Offset(1, 0).Resize(data.Rows.Count - 1, 1)

    CompareCore rngA, rngB, ColumnLabel(data, mCmpA), ColumnLabel(data, mCmpB)
    Exit Sub
Fail:
    MsgBox "Could not compare." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Compare Two Columns"
End Sub

'--- the comparison itself : takes two ranges and puts four result columns to the right ---
Private Sub CompareCore(ByVal rngA As Range, ByVal rngB As Range, _
                     ByVal aName As String, ByVal bName As String)
    Dim ws As Worksheet
    Dim dicA As Object, dicB As Object
    Dim arrA As Variant, arrB As Variant
    Dim outArr() As Variant
    Dim key As String
    Dim r As Long
    Dim nA As Long, nB As Long, maxVal As Long
    Dim aCount As Long, bCount As Long
    Dim aMissing As Long, bMissing As Long, bothCount As Long
    Dim pasteCol As Long, firstRow As Long
    Dim rightPart As Long
    Dim inserted As Boolean, wasFound As Boolean
    Dim aColLetter As String, bColLetter As String
    Dim aAbs As String, bAbs As String
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = rngA.Worksheet
    nA = rngA.Rows.Count
    nB = rngB.Rows.Count
    maxVal = nA
    If nB > maxVal Then maxVal = nB

    If maxVal > 100000 Then
        If MsgBox(Format$(maxVal, "#,##0") & " rows will be compared. This may take a while." & vbCrLf & _
                  "Do you want to continue?", vbOKCancel + vbQuestion, "Compare Two Columns") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    arrA = ColumnToArray(rngA)
    arrB = ColumnToArray(rngB)

    Set dicA = CreateObject("Scripting.Dictionary")
    dicA.CompareMode = 1
    Set dicB = CreateObject("Scripting.Dictionary")
    dicB.CompareMode = 1

    For r = 1 To nA
        key = Trim$(CStr(arrA(r)))
        If Len(key) > 0 Then
            aCount = aCount + 1
            If dicA.Exists(key) Then
                dicA(key) = dicA(key) + 1
            Else
                dicA.Add key, 1
            End If
        End If
    Next r

    For r = 1 To nB
        key = Trim$(CStr(arrB(r)))
        If Len(key) > 0 Then
            bCount = bCount + 1
            If dicB.Exists(key) Then
                dicB(key) = dicB(key) + 1
            Else
                dicB.Add key, 1
            End If
        End If
    Next r

    If aCount = 0 Or bCount = 0 Then
        Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
        MsgBox "There are no values to compare." & vbCrLf & _
               aName & " : " & aCount & " rows / " & bName & " : " & bCount & " rows", _
               vbExclamation, "Compare Two Columns"
        Exit Sub
    End If

    ' where the results go : one cell past the right-hand end of the two ranges
    rightPart = rngA.Column
    If rngB.Column > rightPart Then rightPart = rngB.Column
    pasteCol = rightPart + 2
    firstRow = rngA.Row
    If rngB.Row < firstRow Then firstRow = rngB.Row

    ' build the address strings up front (calling .Address per row is very slow)
    aColLetter = ColumnLetter(rngA.Column)
    bColLetter = ColumnLetter(rngB.Column)
    aAbs = rngA.Address(True, True)
    bAbs = rngB.Address(True, True)

    ReDim outArr(1 To maxVal + 1, 1 To 4)
    outArr(1, 1) = bName & "InOther"
    outArr(1, 2) = bName & "Count"
    outArr(1, 3) = aName & "InOther"
    outArr(1, 4) = aName & "Count"

    For r = 1 To maxVal
        If r <= nA Then
            key = Trim$(CStr(arrA(r)))
            If Len(key) > 0 Then
                If mCmpFormula Then
                    outArr(r + 1, 1) = "=IF(COUNTIF(" & bAbs & "," & aColLetter & (rngA.Row + r - 1) & ")>0,1,0)"
                    outArr(r + 1, 2) = "=COUNTIF(" & bAbs & "," & aColLetter & (rngA.Row + r - 1) & ")"
                Else
                    outArr(r + 1, 1) = IIf(dicB.Exists(key), 1, 0)
                    outArr(r + 1, 2) = IIf(dicB.Exists(key), dicB(key), 0)
                End If
                If dicB.Exists(key) Then
                    bothCount = bothCount + 1
                Else
                    aMissing = aMissing + 1
                End If
            End If
        End If

        If r <= nB Then
            key = Trim$(CStr(arrB(r)))
            If Len(key) > 0 Then
                If mCmpFormula Then
                    outArr(r + 1, 3) = "=IF(COUNTIF(" & aAbs & "," & bColLetter & (rngB.Row + r - 1) & ")>0,1,0)"
                    outArr(r + 1, 4) = "=COUNTIF(" & aAbs & "," & bColLetter & (rngB.Row + r - 1) & ")"
                Else
                    outArr(r + 1, 3) = IIf(dicA.Exists(key), 1, 0)
                    outArr(r + 1, 4) = IIf(dicA.Exists(key), dicA(key), 0)
                End If
                If Not dicA.Exists(key) Then bMissing = bMissing + 1
            End If
        End If
    Next r

    ' when there is data to the right, ask, then insert four columns
    inserted = False
    On Error Resume Next
    wasFound = (Application.CountA(ws.Range(ws.Cells(firstRow - 1, pasteCol), _
                                        ws.Cells(firstRow + maxVal - 1, pasteCol + 3))) > 0)
    Err.Clear
    On Error GoTo Fail

    If wasFound Then
        Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
        If MsgBox("There is already data where the result would go." & vbCrLf & _
                  "Insert 4 empty columns and fill?" & vbCrLf & vbCrLf & _
                  "Existing data will shift to the right." & vbCrLf & _
                  "You can undo this with Ctrl+Z.", _
                  vbOKCancel + vbQuestion, "Compare Two Columns") = vbCancel Then Exit Sub
        Application.ScreenUpdating = False
        ws.Range(ws.Columns(pasteCol), ws.Columns(pasteCol + 3)).Insert Shift:=xlToRight
        inserted = True
        mInsWb = ws.Parent.Name
        mInsWs = ws.Name
        mInsCol = pasteCol
        mInsCount = 4
        mInsOn = True
    Else
        SaveCriteriaBlock ws.Range(ws.Cells(firstRow - 1, pasteCol), ws.Cells(firstRow + maxVal - 1, pasteCol + 3))
        mInsOn = False
    End If

    ws.Range(ws.Cells(firstRow - 1, pasteCol), _
             ws.Cells(firstRow + maxVal - 1, pasteCol + 3)).Formula = outArr

    With ws.Range(ws.Cells(firstRow - 1, pasteCol), ws.Cells(firstRow - 1, pasteCol + 3))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Columns(pasteCol), ws.Columns(pasteCol + 3)).AutoFit

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    If inserted Then
        Application.OnUndo "Undo Compare Two Columns", MacroPrefix() & "UndoInsertedColumns"
    ElseIf mCondOn Then
        Application.OnUndo "Undo Compare Two Columns", MacroPrefix() & "UndoCriteriaBlock"
    End If
    Err.Clear
    On Error GoTo Fail

    MsgBox "Comparison finished." & vbCrLf & vbCrLf & _
           "In both lists : " & Format$(bothCount, "#,##0") & " rows" & vbCrLf & vbCrLf & _
           aName & " " & Format$(aCount, "#,##0") & " of " & bName & " missing from : " & _
           Format$(aMissing, "#,##0") & " rows" & vbCrLf & _
           bName & " " & Format$(bCount, "#,##0") & " of " & aName & " missing from : " & _
           Format$(bMissing, "#,##0") & " rows" & vbCrLf & vbCrLf & _
           "Inserted as : " & IIf(mCmpFormula, "formulas (COUNTIF)", "Value") & vbCrLf & _
           "Filter on 0 to see only the ones that are missing." & vbCrLf & _
           "You can undo this with Ctrl+Z.", vbInformation, "Compare Two Columns"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not compare." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Compare Two Columns"
End Sub

'--- a single-column range into a one-dimensional array ---
Private Function ColumnToArray(ByVal rg As Range) As Variant
    Dim v As Variant
    Dim a() As Variant
    Dim r As Long

    If rg.Rows.Count = 1 Then
        ReDim a(1 To 1)
        a(1) = rg.Cells(1, 1).Value
        ColumnToArray = a
        Exit Function
    End If

    v = rg.Value
    ReDim a(1 To rg.Rows.Count)
    For r = 1 To rg.Rows.Count
        a(r) = v(r, 1)
    Next r
    ColumnToArray = a
End Function

'============================================================
'  test one shortcut
'    narrows down why a particular shortcut is not working.
'    binds the key to a plain message for a moment, then puts it back.
'============================================================


'--- compare-two-columns toggle wrappers ---
Public Sub CmpAToggle01()
    CmpAToggle 1
End Sub
Public Sub CmpAToggle02()
    CmpAToggle 2
End Sub
Public Sub CmpAToggle03()
    CmpAToggle 3
End Sub
Public Sub CmpAToggle04()
    CmpAToggle 4
End Sub
Public Sub CmpAToggle05()
    CmpAToggle 5
End Sub
Public Sub CmpAToggle06()
    CmpAToggle 6
End Sub
Public Sub CmpAToggle07()
    CmpAToggle 7
End Sub
Public Sub CmpAToggle08()
    CmpAToggle 8
End Sub
Public Sub CmpAToggle09()
    CmpAToggle 9
End Sub
Public Sub CmpAToggle10()
    CmpAToggle 10
End Sub
Public Sub CmpAToggle11()
    CmpAToggle 11
End Sub
Public Sub CmpAToggle12()
    CmpAToggle 12
End Sub
Public Sub CmpAToggle13()
    CmpAToggle 13
End Sub
Public Sub CmpAToggle14()
    CmpAToggle 14
End Sub
Public Sub CmpAToggle15()
    CmpAToggle 15
End Sub
Public Sub CmpAToggle16()
    CmpAToggle 16
End Sub
Public Sub CmpAToggle17()
    CmpAToggle 17
End Sub
Public Sub CmpAToggle18()
    CmpAToggle 18
End Sub
Public Sub CmpAToggle19()
    CmpAToggle 19
End Sub
Public Sub CmpAToggle20()
    CmpAToggle 20
End Sub
Public Sub CmpAToggle21()
    CmpAToggle 21
End Sub
Public Sub CmpAToggle22()
    CmpAToggle 22
End Sub
Public Sub CmpAToggle23()
    CmpAToggle 23
End Sub
Public Sub CmpAToggle24()
    CmpAToggle 24
End Sub
Public Sub CmpAToggle25()
    CmpAToggle 25
End Sub
Public Sub CmpAToggle26()
    CmpAToggle 26
End Sub
Public Sub CmpAToggle27()
    CmpAToggle 27
End Sub
Public Sub CmpAToggle28()
    CmpAToggle 28
End Sub
Public Sub CmpAToggle29()
    CmpAToggle 29
End Sub
Public Sub CmpAToggle30()
    CmpAToggle 30
End Sub
Public Sub CmpAToggle31()
    CmpAToggle 31
End Sub
Public Sub CmpAToggle32()
    CmpAToggle 32
End Sub
Public Sub CmpAToggle33()
    CmpAToggle 33
End Sub
Public Sub CmpAToggle34()
    CmpAToggle 34
End Sub
Public Sub CmpAToggle35()
    CmpAToggle 35
End Sub
Public Sub CmpAToggle36()
    CmpAToggle 36
End Sub
Public Sub CmpAToggle37()
    CmpAToggle 37
End Sub
Public Sub CmpAToggle38()
    CmpAToggle 38
End Sub
Public Sub CmpAToggle39()
    CmpAToggle 39
End Sub
Public Sub CmpAToggle40()
    CmpAToggle 40
End Sub
Public Sub CmpAToggle41()
    CmpAToggle 41
End Sub
Public Sub CmpAToggle42()
    CmpAToggle 42
End Sub
Public Sub CmpAToggle43()
    CmpAToggle 43
End Sub
Public Sub CmpAToggle44()
    CmpAToggle 44
End Sub
Public Sub CmpAToggle45()
    CmpAToggle 45
End Sub
Public Sub CmpAToggle46()
    CmpAToggle 46
End Sub
Public Sub CmpAToggle47()
    CmpAToggle 47
End Sub
Public Sub CmpAToggle48()
    CmpAToggle 48
End Sub
Public Sub CmpAToggle49()
    CmpAToggle 49
End Sub
Public Sub CmpAToggle50()
    CmpAToggle 50
End Sub
Public Sub CmpAToggle51()
    CmpAToggle 51
End Sub
Public Sub CmpAToggle52()
    CmpAToggle 52
End Sub
Public Sub CmpAToggle53()
    CmpAToggle 53
End Sub
Public Sub CmpAToggle54()
    CmpAToggle 54
End Sub
Public Sub CmpAToggle55()
    CmpAToggle 55
End Sub
Public Sub CmpAToggle56()
    CmpAToggle 56
End Sub
Public Sub CmpAToggle57()
    CmpAToggle 57
End Sub
Public Sub CmpAToggle58()
    CmpAToggle 58
End Sub
Public Sub CmpAToggle59()
    CmpAToggle 59
End Sub
Public Sub CmpAToggle60()
    CmpAToggle 60
End Sub
Public Sub CmpBToggle01()
    CmpBToggle 1
End Sub
Public Sub CmpBToggle02()
    CmpBToggle 2
End Sub
Public Sub CmpBToggle03()
    CmpBToggle 3
End Sub
Public Sub CmpBToggle04()
    CmpBToggle 4
End Sub
Public Sub CmpBToggle05()
    CmpBToggle 5
End Sub
Public Sub CmpBToggle06()
    CmpBToggle 6
End Sub
Public Sub CmpBToggle07()
    CmpBToggle 7
End Sub
Public Sub CmpBToggle08()
    CmpBToggle 8
End Sub
Public Sub CmpBToggle09()
    CmpBToggle 9
End Sub
Public Sub CmpBToggle10()
    CmpBToggle 10
End Sub
Public Sub CmpBToggle11()
    CmpBToggle 11
End Sub
Public Sub CmpBToggle12()
    CmpBToggle 12
End Sub
Public Sub CmpBToggle13()
    CmpBToggle 13
End Sub
Public Sub CmpBToggle14()
    CmpBToggle 14
End Sub
Public Sub CmpBToggle15()
    CmpBToggle 15
End Sub
Public Sub CmpBToggle16()
    CmpBToggle 16
End Sub
Public Sub CmpBToggle17()
    CmpBToggle 17
End Sub
Public Sub CmpBToggle18()
    CmpBToggle 18
End Sub
Public Sub CmpBToggle19()
    CmpBToggle 19
End Sub
Public Sub CmpBToggle20()
    CmpBToggle 20
End Sub
Public Sub CmpBToggle21()
    CmpBToggle 21
End Sub
Public Sub CmpBToggle22()
    CmpBToggle 22
End Sub
Public Sub CmpBToggle23()
    CmpBToggle 23
End Sub
Public Sub CmpBToggle24()
    CmpBToggle 24
End Sub
Public Sub CmpBToggle25()
    CmpBToggle 25
End Sub
Public Sub CmpBToggle26()
    CmpBToggle 26
End Sub
Public Sub CmpBToggle27()
    CmpBToggle 27
End Sub
Public Sub CmpBToggle28()
    CmpBToggle 28
End Sub
Public Sub CmpBToggle29()
    CmpBToggle 29
End Sub
Public Sub CmpBToggle30()
    CmpBToggle 30
End Sub
Public Sub CmpBToggle31()
    CmpBToggle 31
End Sub
Public Sub CmpBToggle32()
    CmpBToggle 32
End Sub
Public Sub CmpBToggle33()
    CmpBToggle 33
End Sub
Public Sub CmpBToggle34()
    CmpBToggle 34
End Sub
Public Sub CmpBToggle35()
    CmpBToggle 35
End Sub
Public Sub CmpBToggle36()
    CmpBToggle 36
End Sub
Public Sub CmpBToggle37()
    CmpBToggle 37
End Sub
Public Sub CmpBToggle38()
    CmpBToggle 38
End Sub
Public Sub CmpBToggle39()
    CmpBToggle 39
End Sub
Public Sub CmpBToggle40()
    CmpBToggle 40
End Sub
Public Sub CmpBToggle41()
    CmpBToggle 41
End Sub
Public Sub CmpBToggle42()
    CmpBToggle 42
End Sub
Public Sub CmpBToggle43()
    CmpBToggle 43
End Sub
Public Sub CmpBToggle44()
    CmpBToggle 44
End Sub
Public Sub CmpBToggle45()
    CmpBToggle 45
End Sub
Public Sub CmpBToggle46()
    CmpBToggle 46
End Sub
Public Sub CmpBToggle47()
    CmpBToggle 47
End Sub
Public Sub CmpBToggle48()
    CmpBToggle 48
End Sub
Public Sub CmpBToggle49()
    CmpBToggle 49
End Sub
Public Sub CmpBToggle50()
    CmpBToggle 50
End Sub
Public Sub CmpBToggle51()
    CmpBToggle 51
End Sub
Public Sub CmpBToggle52()
    CmpBToggle 52
End Sub
Public Sub CmpBToggle53()
    CmpBToggle 53
End Sub
Public Sub CmpBToggle54()
    CmpBToggle 54
End Sub
Public Sub CmpBToggle55()
    CmpBToggle 55
End Sub
Public Sub CmpBToggle56()
    CmpBToggle 56
End Sub
Public Sub CmpBToggle57()
    CmpBToggle 57
End Sub
Public Sub CmpBToggle58()
    CmpBToggle 58
End Sub
Public Sub CmpBToggle59()
    CmpBToggle 59
End Sub
Public Sub CmpBToggle60()
    CmpBToggle 60
End Sub

'============================================================
'  auto fill series (Ctrl+Alt+Q)
'    reads patterns such as 1,2,3 / Mon,Tue,Wed / dates and extends them to the end of the table.
'    at least two seed cells are needed to read a pattern. (same as Excel's own fill handle)
'============================================================
Public Sub AutoFillSeries()
    Dim ws As Worksheet
    Dim sel As Range
    Dim blk As Range
    Dim target As Range
    Dim c As Range
    Dim endRow As Long
    Dim overwriteCount As Long
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Auto Fill Series"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the seed cells, then run again.", vbExclamation, "Auto Fill Series"
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "Separate ranges cannot be processed at once.", _
               vbExclamation, "Auto Fill Series"
        Exit Sub
    End If

    On Error GoTo Fail

    Set blk = SkipNonHeaderRows(sel.Cells(1, 1).CurrentRegion)
    If blk Is Nothing Then
        MsgBox "Could not find a table.", vbExclamation, "Auto Fill Series"
        Exit Sub
    End If

    endRow = blk.Row + blk.Rows.Count - 1
    If endRow <= sel.Row + sel.Rows.Count - 1 Then
        MsgBox "There are no rows below to fill." & vbCrLf & _
               "The cursor is on the last row, or no table was found.", _
               vbExclamation, "Auto Fill Series"
        Exit Sub
    End If

    Set target = ws.Range(ws.Cells(sel.Row, sel.Column), _
                        ws.Cells(endRow, sel.Column + sel.Columns.Count - 1))

    '--- confirm when there is already data below the seed ---
    overwriteCount = 0
    On Error Resume Next
    For Each c In sel.Columns
        overwriteCount = overwriteCount + Application.CountA(ws.Range( _
                    ws.Cells(sel.Row + sel.Rows.Count, c.Column), ws.Cells(endRow, c.Column)))
    Next c
    Err.Clear
    On Error GoTo Fail

    If overwriteCount > 0 Then
        If MsgBox("There are already values below in " & Format$(overwriteCount, "#,##0") & " cells." & vbCrLf & _
                  "Overwrite?" & vbCrLf & vbCrLf & _
                  "Fill range : " & (sel.Row + sel.Rows.Count) & " to row " & endRow & " rows" & vbCrLf & _
                  "You can undo this with Ctrl+Z.", _
                  vbOKCancel + vbExclamation, "Auto Fill Series") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    SaveCriteriaBlock target

    sel.AutoFill Destination:=target, Type:=xlFillDefault

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "Undo Auto Fill Series", MacroPrefix() & "UndoCriteriaBlock"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "Continuing the pattern, " & (sel.Row + sel.Rows.Count) & " to row " & endRow & _
                            " rows filled." & _
                            IIf(sel.Rows.Count < 2, "   (only one seed cell, so the value is copied)", "") & _
                            IIf(mCondOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not fill." & vbCrLf & _
           "Error " & errNo & ": " & errDesc & vbCrLf & vbCrLf & _
           "At least two seed cells are needed to read a pattern.", _
           vbExclamation, "Auto Fill Series"
End Sub

'--- removing rows above the header : kept so it can be undone ---
Private Sub SaveRowsAboveHeader(ByVal ws As Worksheet, ByVal rowCount2 As Long)
    Dim r As Long, c As Long, nC As Long
    Dim vF() As Variant, vN() As Variant, vC() As Variant
    Dim vI() As Variant, vB() As Variant, vA() As Variant, vH() As Variant

    mDelOn = False
    On Error GoTo Fail

    nC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If nC < 1 Then Exit Sub
    If nC > 200 Then nC = 200
    If rowCount2 * nC > 20000 Then Exit Sub

    ReDim vF(1 To rowCount2, 1 To nC)
    ReDim vN(1 To rowCount2, 1 To nC)
    ReDim vC(1 To rowCount2, 1 To nC)
    ReDim vI(1 To rowCount2, 1 To nC)
    ReDim vB(1 To rowCount2, 1 To nC)
    ReDim vA(1 To rowCount2, 1 To nC)
    ReDim vH(1 To rowCount2)

    For r = 1 To rowCount2
        vH(r) = ws.Rows(r).RowHeight
        For c = 1 To nC
            With ws.Cells(r, c)
                vF(r, c) = .Formula
                vN(r, c) = .NumberFormat
                vC(r, c) = .Interior.Color
                vI(r, c) = .Interior.ColorIndex
                vB(r, c) = .Font.Bold
                vA(r, c) = .HorizontalAlignment
            End With
        Next c
    Next r

    mDelWb = ws.Parent.Name
    mDelWs = ws.Name
    mDelRows = rowCount2
    mDelCols = nC
    mDelF = vF
    mDelN = vN
    mDelC = vC
    mDelI = vI
    mDelB = vB
    mDelA = vA
    mDelH = vH
    mDelOn = True
    Exit Sub
Fail:
    mDelOn = False
End Sub

'------------------------------------------------------------
' undo removing the rows above the header (called by Ctrl+Z)
'------------------------------------------------------------
Public Sub UndoRemoveRowsAboveHeader()
    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim oldScreen As Boolean

    If Not mDelOn Or mDelRows < 1 Then
        MsgBox "There are no removed rows to undo.", vbInformation, "Remove Rows Above Header"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False

    Set ws = Application.Workbooks(mDelWb).Worksheets(mDelWs)
    ws.Range(ws.Rows(1), ws.Rows(mDelRows)).Insert Shift:=xlDown

    For r = 1 To mDelRows
        ws.Rows(r).RowHeight = mDelH(r)
        For c = 1 To mDelCols
            With ws.Cells(r, c)
                .NumberFormat = mDelN(r, c)
                .Formula = mDelF(r, c)
                If mDelI(r, c) = xlColorIndexNone Then
                    .Interior.ColorIndex = xlColorIndexNone
                Else
                    .Interior.Color = mDelC(r, c)
                End If
                .Font.Bold = mDelB(r, c)
                .HorizontalAlignment = mDelA(r, c)
            End With
        Next c
    Next r

    mDelOn = False
    mDelRows = 0
    Err.Clear
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    Application.StatusBar = "The removed rows were restored."
    ScheduleStatusClear
End Sub

'============================================================
'  flash fill (Ctrl+Alt+A)  -  an upgrade on Ctrl+E
'    takes the cells above the cursor as the example and fills from the cursor to the end of the table.
'    unlike Ctrl+E, the range to fill does not have to be selected.
'============================================================
Public Sub FlashFillToEnd()
    Dim ws As Worksheet
    Dim cursorCell As Range
    Dim blk As Range
    Dim target As Range
    Dim sampleRows As Long, startRow As Long, endRow As Long
    Dim colIdx As Long
    Dim filledCount As Long, r As Long
    Dim overwriteCount As Long
    Dim noticeText As String
    Dim filledRange As Range
    Dim resultVal As Variant
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Flash Fill"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select the cell to start filling from, then run again.", vbExclamation, "Flash Fill"
        Exit Sub
    End If

    Set cursorCell = ActiveCell
    colIdx = cursorCell.Column
    startRow = cursorCell.Row

    On Error GoTo Fail

    Set blk = SkipNonHeaderRows(cursorCell.CurrentRegion)
    If blk Is Nothing Then
        MsgBox "Could not find a table." & vbCrLf & _
               "The rule can only be read when the next column has data.", _
               vbExclamation, "Flash Fill"
        Exit Sub
    End If

    endRow = blk.Row + blk.Rows.Count - 1
    If endRow < startRow Then
        MsgBox "There are no rows to fill.", vbExclamation, "Flash Fill"
        Exit Sub
    End If

    '--- how many example cells sit above the cursor ---
    sampleRows = 0
    For r = startRow - 1 To blk.Row Step -1
        If Len(CStr(ws.Cells(r, colIdx).Value)) = 0 Then Exit For
        sampleRows = sampleRows + 1
    Next r

    If sampleRows = 0 Then
        MsgBox "There is no example in the cell just above." & vbCrLf & vbCrLf & _
               "Type the result you want in one or two cells, then" & vbCrLf & _
               "Put the cursor on the cell below it and run again.", _
               vbExclamation, "Flash Fill"
        Exit Sub
    End If

    '--- whether the area to fill already holds data ---
    overwriteCount = 0
    On Error Resume Next
    overwriteCount = Application.CountA(ws.Range(ws.Cells(startRow, colIdx), ws.Cells(endRow, colIdx)))
    Err.Clear
    On Error GoTo Fail

    If sampleRows = 1 Or overwriteCount > 0 Then
        noticeText = ""
        If sampleRows = 1 Then
            noticeText = noticeText & "· Only one example cell. Two or more reads the rule more reliably." & vbCrLf
        End If
        If overwriteCount > 0 Then
            noticeText = noticeText & "· There are already values where it would fill : " & Format$(overwriteCount, "#,##0") & " cells already hold data. They will be overwritten." & vbCrLf
        End If
        If MsgBox(noticeText & vbCrLf & _
                  "Fill range : " & startRow & " to row " & endRow & " rows" & vbCrLf & _
                  "You can undo this with Ctrl+Z." & vbCrLf & vbCrLf & _
                  "Do you want to continue?", vbOKCancel + vbQuestion, "Flash Fill") = vbCancel Then Exit Sub
    End If

    Set target = ws.Range(ws.Cells(startRow - sampleRows, colIdx), ws.Cells(endRow, colIdx))

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    SaveCriteriaBlock ws.Range(ws.Cells(startRow, colIdx), ws.Cells(endRow, colIdx))

    On Error Resume Next
    target.FlashFill
    Err.Clear
    On Error GoTo Fail

    ' count whether anything was actually filled
    For r = startRow To endRow
        If Len(CStr(ws.Cells(r, colIdx).Value)) > 0 Then filledCount = filledCount + 1
    Next r

    ' FlashFill is built into Excel, so it takes the undo stack for itself.
    ' rewriting the filled values from VBA takes the stack back.
    If filledCount > 0 Then
        Set filledRange = ws.Range(ws.Cells(startRow, colIdx), ws.Cells(endRow, colIdx))
        resultVal = filledRange.Value
        filledRange.Value = resultVal
    End If

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    If filledCount = 0 Then
        MsgBox "Could not read the rule." & vbCrLf & vbCrLf & _
               "· Try giving two or more example cells." & vbCrLf & _
               "· The source data must be in the column next to it." & vbCrLf & _
               "· Flash Fill does nothing when the rule is unclear.", _
               vbExclamation, "Flash Fill"
        Exit Sub
    End If

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "Undo Flash Fill", MacroPrefix() & "UndoCriteriaBlock"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "Example " & sampleRows & "Reading " & startRow & " to row " & endRow & _
                            " rows filled.  (" & Format$(filledCount, "#,##0") & " cells)" & _
                            IIf(mCondOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not fill." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Flash Fill"
End Sub

'============================================================
'  copy / paste formatting (Ctrl+Alt+X)
'    press once and the cell's formatting is remembered,
'    select somewhere else and press again to apply it.
'    no need to reach for the format painter with the mouse.
'============================================================
Public Sub FormatBrush()
    Dim ws As Worksheet
    Dim rng As Range
    Dim source As Range
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Please run this on a worksheet.", vbExclamation, "Format Brush"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell first.", vbExclamation, "Format Brush"
        Exit Sub
    End If
    Set rng = Selection

    On Error GoTo Fail

    '--- nothing remembered yet : copy ---
    If Not mFmtSrcOn Then
        mFmtSrcWb = ws.Parent.Name
        mFmtSrcWs = ws.Name
        mFmtSrcAddr = rng.Address
        mFmtSrcOn = True

        ' Esc cancels it, and it clears itself after 40 seconds
        On Error Resume Next
        Application.OnKey "{ESC}", MacroPrefix() & "CancelFormatBrush"
        mFmtSrcTime = Now + TimeSerial(0, 0, FORMAT_BRUSH_SECONDS)
        Application.OnTime mFmtSrcTime, MacroPrefix() & "FormatBrushTimeout"
        Err.Clear
        On Error GoTo Fail

        ' this message is not cleared after 4 seconds - it stays until the paste
        Application.StatusBar = "[ Format brush ready ]  " & rng.Address(False, False) & _
                                "   ->  select where to paste, then Ctrl+Alt+X" & _
                                "      ( Esc to cancel  " & FORMAT_BRUSH_SECONDS & " s until it clears )"
        Exit Sub
    End If

    '--- something remembered : paste ---
    Set source = Nothing
    On Error Resume Next
    Set source = Application.Workbooks(mFmtSrcWb).Worksheets(mFmtSrcWs).Range(mFmtSrcAddr)
    Err.Clear
    On Error GoTo Fail

    If source Is Nothing Then
        mFmtSrcOn = False
        MsgBox "No remembered format was found." & vbCrLf & _
               "Select what to copy from and run again.", vbExclamation, "Format Brush"
        Exit Sub
    End If

    ' pressing again in the same spot cancels it
    If StrComp(mFmtSrcWs, ws.Name, vbTextCompare) = 0 And _
       StrComp(mFmtSrcAddr, rng.Address, vbTextCompare) = 0 Then
        ClearFormatBrush
        Application.StatusBar = "Format brush cancelled."
        ScheduleStatusClear
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    SaveCriteriaBlock rng

    source.Copy
    rng.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ClearFormatBrush
    Application.ScreenUpdating = oldScreen

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "Undo Paste Format", MacroPrefix() & "UndoCriteriaBlock"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "Format pasted. (" & rng.Address(False, False) & ")" & _
                            IIf(mCondOn, "  |  Ctrl+Z to undo", "")
    ScheduleStatusClear
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
    ClearFormatBrush
    MsgBox "Could not handle the format." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Format Brush"
End Sub

'============================================================
'  compare the same rows (from the Ctrl+Alt+D popup)
'    puts two columns side by side row by row and marks only the rows that differ.
'    e.g. Team vs New Team  ->  picks out only the people who are moving.
'============================================================
Private Sub RowCompareDirect(ByVal rngA As Range, ByVal rngB As Range)
    Dim a As Range, b As Range

    Set a = rngA
    Set b = rngB
    TrimCompareRange a, b
    Set rngA = a
    Set rngB = b

    If rngA.Columns.Count > 1 Or rngB.Columns.Count > 1 Then
        MsgBox "Please select one column at a time." & vbCrLf & _
               "Hold Ctrl and click the two columns you want to compare.", _
               vbExclamation, "Compare Within Each Row"
        Exit Sub
    End If
    RowCompareCore rngA, rngB, RangeHeader(rngA), RangeHeader(rngB)
End Sub

Public Sub RunRowCompare()
    Dim ws As Worksheet
    Dim data As Range
    Dim rngA As Range, rngB As Range

    On Error GoTo Fail
    Set ws = ActiveSheet
    Set data = ws.Range(mCmpAddr)

    Set rngA = data.Columns(mCmpA).Offset(1, 0).Resize(data.Rows.Count - 1, 1)
    Set rngB = data.Columns(mCmpB).Offset(1, 0).Resize(data.Rows.Count - 1, 1)

    RowCompareCore rngA, rngB, ColumnLabel(data, mCmpA), ColumnLabel(data, mCmpB)
    Exit Sub
Fail:
    MsgBox "Could not compare." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Compare Within Each Row"
End Sub

'--- the comparison itself ---
Private Sub RowCompareCore(ByVal rngA As Range, ByVal rngB As Range, _
                       ByVal aName As String, ByVal bName As String)
    Dim ws As Worksheet
    Dim arrA As Variant, arrB As Variant
    Dim result() As Variant
    Dim r As Long, n As Long
    Dim vA As String, vB As String
    Dim diffCount As Long, blankCell As Long
    Dim pasteCol As Long, firstRow As Long
    Dim resultRange As Range, colorRange As Range
    Dim painted As Range
    Dim inserted As Boolean, wasFound As Boolean
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail
    Set ws = rngA.Worksheet

    n = rngA.Rows.Count
    If rngB.Rows.Count < n Then n = rngB.Rows.Count
    If rngA.Rows.Count <> rngB.Rows.Count Then
        If MsgBox("The two ranges have different row counts." & vbCrLf & vbCrLf & _
                  aName & " : " & rngA.Rows.Count & " rows" & vbCrLf & _
                  bName & " : " & rngB.Rows.Count & " rows" & vbCrLf & vbCrLf & _
                  "the shorter side, " & n & " rows only will be compared. Continue?", _
                  vbOKCancel + vbExclamation, "Compare Within Each Row") = vbCancel Then Exit Sub
    End If

    If n < 1 Then
        MsgBox "There are no rows to compare.", vbExclamation, "Compare Within Each Row"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    arrA = ColumnToArray(rngA)
    arrB = ColumnToArray(rngB)

    ReDim result(1 To n + 1, 1 To 1)
    result(1, 1) = aName & " ≠ " & bName

    For r = 1 To n
        vA = Trim$(CStr(arrA(r)))
        vB = Trim$(CStr(arrB(r)))

        If Len(vA) = 0 And Len(vB) = 0 Then
            result(r + 1, 1) = 0
        ElseIf StrComp(vA, vB, vbTextCompare) = 0 Then
            result(r + 1, 1) = 0
        Else
            result(r + 1, 1) = 1
            diffCount = diffCount + 1
            If Len(vA) = 0 Or Len(vB) = 0 Then blankCell = blankCell + 1
            If painted Is Nothing Then
                Set painted = Application.Union(rngA.Cells(r, 1), rngB.Cells(r, 1))
            Else
                Set painted = Application.Union(painted, rngA.Cells(r, 1), rngB.Cells(r, 1))
            End If
        End If
    Next r

    ' where the result columns go : one cell past the right-hand end of the two ranges
    pasteCol = rngA.Column
    If rngB.Column > pasteCol Then pasteCol = rngB.Column
    pasteCol = pasteCol + 2
    firstRow = rngA.Row
    If rngB.Row < firstRow Then firstRow = rngB.Row

    ' when there is data to the right, ask, then insert the columns
    inserted = False
    On Error Resume Next
    wasFound = (Application.CountA(ws.Range(ws.Cells(firstRow - 1, pasteCol), _
                                        ws.Cells(firstRow + n - 1, pasteCol))) > 0)
    Err.Clear
    On Error GoTo Fail

    If wasFound Then
        Application.ScreenUpdating = oldScreen
        If MsgBox("There is already data where the result would go." & vbCrLf & _
                  "Insert one empty column and fill?" & vbCrLf & vbCrLf & _
                  "Existing data will shift to the right." & vbCrLf & _
                  "You can undo this with Ctrl+Z.", _
                  vbOKCancel + vbQuestion, "Compare Within Each Row") = vbCancel Then Exit Sub
        Application.ScreenUpdating = False
        ws.Columns(pasteCol).Insert Shift:=xlToRight
        inserted = True
    End If

    Set resultRange = ws.Range(ws.Cells(firstRow - 1, pasteCol), ws.Cells(firstRow + n - 1, pasteCol))
    Set colorRange = Application.Union(rngA.Resize(n, 1), rngB.Resize(n, 1))

    SaveRowCompare resultRange, colorRange, inserted, pasteCol

    resultRange.Value = result
    With ws.Cells(firstRow - 1, pasteCol)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    resultRange.HorizontalAlignment = xlCenter
    ws.Columns(pasteCol).AutoFit

    If Not painted Is Nothing Then painted.Interior.Color = RGB(255, 199, 206)

    Application.ScreenUpdating = oldScreen

    If mDifOn Then
        On Error Resume Next
        Application.OnUndo "Undo Compare Within Each Row", MacroPrefix() & "UndoRowCompare"
        Err.Clear
        On Error GoTo Fail
    End If

    MsgBox "Comparison finished." & vbCrLf & vbCrLf & _
           aName & " vs " & bName & vbCrLf & vbCrLf & _
           "Rows compared : " & Format$(n, "#,##0") & " rows" & vbCrLf & _
           "Rows that differ : " & Format$(diffCount, "#,##0") & " rows" & _
           IIf(blankCell > 0, "   (one side had " & blankCell & " rows included)", "") & vbCrLf & vbCrLf & _
           "Rows that differ are colored pink." & vbCrLf & _
           "Filter the result column on 1 to see them together." & vbCrLf & _
           "You can undo this with Ctrl+Z.", vbInformation, "Compare Within Each Row"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "Could not compare." & vbCrLf & _
           "Error " & errNo & ": " & errDesc, vbExclamation, "Compare Within Each Row"
End Sub

'--- kept so it can be undone : the result column values and the compared ranges' colors ---
Private Sub SaveRowCompare(ByVal resultRange As Range, ByVal colorRange As Range, _
                       ByVal inserted As Boolean, ByVal pasteCol As Long)
    Dim ar As Range, cel As Range
    Dim v() As Variant, vC() As Variant, vI() As Variant
    Dim i As Long, n As Long

    mDifOn = False
    On Error GoTo Fail

    If resultRange.Rows.Count > 60000 Then Exit Sub

    ReDim v(1 To resultRange.Rows.Count)
    For i = 1 To resultRange.Rows.Count
        v(i) = resultRange.Cells(i, 1).Formula
    Next i

    n = 0
    For Each ar In colorRange.Areas
        n = n + ar.Cells.Count
    Next ar
    If n > 120000 Then Exit Sub

    ReDim vC(1 To n)
    ReDim vI(1 To n)
    i = 0
    For Each ar In colorRange.Areas
        For Each cel In ar.Cells
            i = i + 1
            vC(i) = cel.Interior.Color
            vI(i) = cel.Interior.ColorIndex
        Next cel
    Next ar

    mDifWb = resultRange.Worksheet.Parent.Name
    mDifWs = resultRange.Worksheet.Name
    mDifResAddr = resultRange.Address
    mDifColAddr = colorRange.Address
    mDifRes = v
    mDifColor = vC
    mDifIndex = vI
    mDifIns = inserted
    mDifInsCol = pasteCol
    mDifOn = True
    Exit Sub
Fail:
    mDifOn = False
End Sub

'------------------------------------------------------------
' undo comparing the same rows (called by Ctrl+Z)
'------------------------------------------------------------
Public Sub UndoRowCompare()
    Dim ws As Worksheet
    Dim resultRange As Range, colorRange As Range
    Dim ar As Range, cel As Range
    Dim i As Long
    Dim oldScreen As Boolean

    If Not mDifOn Or Len(mDifResAddr) = 0 Then
        MsgBox "There is no comparison to undo.", vbInformation, "Compare Within Each Row"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False

    Set ws = Application.Workbooks(mDifWb).Worksheets(mDifWs)
    Set colorRange = ws.Range(mDifColAddr)

    If mDifIns Then
        ws.Columns(mDifInsCol).Delete Shift:=xlToLeft
    Else
        Set resultRange = ws.Range(mDifResAddr)
        For i = 1 To resultRange.Rows.Count
            resultRange.Cells(i, 1).Formula = mDifRes(i)
        Next i
    End If

    i = 0
    For Each ar In colorRange.Areas
        For Each cel In ar.Cells
            i = i + 1
            If mDifIndex(i) = xlColorIndexNone Then
                cel.Interior.ColorIndex = xlColorIndexNone
            Else
                cel.Interior.Color = mDifColor(i)
            End If
        Next cel
    Next ar

    mDifOn = False
    mDifResAddr = ""
    Err.Clear
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0

    Application.StatusBar = "The comparison was undone."
    ScheduleStatusClear
End Sub

'--- two spots picked with Ctrl : a small popup to choose the job ---
Private Sub TwoSpotMenu()
    Dim cb As CommandBar
    Dim btn As CommandBarButton

    On Error Resume Next
    Application.CommandBars(BAR_TWO_SPOT).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=BAR_TWO_SPOT, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    AddMenuTitle cb, "What do you want to do with the two selections"

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Compare between lists   (1/0 both ways, 4 result columns)"
    btn.OnAction = MacroPrefix() & "TwoSpot_ListCompare"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  Compare within each row   (color the differing rows, 1 result column)"
    btn.OnAction = MacroPrefix() & "TwoSpot_RowCompare"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Between lists : is the value anywhere in the other column"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Within each row : compare the two cells side by side"
    btn.Style = msoButtonCaption
    btn.Enabled = False

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "Could not show the menu." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Compare Two Columns"
End Sub

Public Sub TwoSpot_ListCompare()
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = Application.ActiveWorkbook.Worksheets(mCmpDirWs)
    CompareDirect ws.Range(mCmpDirA), ws.Range(mCmpDirB)
    Exit Sub
Fail:
    MsgBox "Could not compare." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Compare Two Columns"
End Sub

Public Sub TwoSpot_RowCompare()
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = Application.ActiveWorkbook.Worksheets(mCmpDirWs)
    RowCompareDirect ws.Range(mCmpDirA), ws.Range(mCmpDirB)
    Exit Sub
Fail:
    MsgBox "Could not compare." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Compare Two Columns"
End Sub

'------------------------------------------------------------
' tidy up the ranges to compare
'   keeps a whole-column pick from dragging in 1,048,576 rows.
'   trims it to below the table's header down to the last data row.
'   lines the two ranges up over the same span so the rows do not slip.
'------------------------------------------------------------
Private Sub TrimCompareRange(ByRef rngA As Range, ByRef rngB As Range)
    Dim ws As Worksheet
    Dim blk As Range
    Dim lastRow As Range
    Dim headerRow As Long, endRow As Long
    Dim overflow As Boolean

    On Error GoTo Done
    Set ws = rngA.Worksheet
    If rngB.Worksheet.Name <> ws.Name Then Exit Sub

    ' tidied up when either side is a whole column or runs past 200,000 rows
    overflow = (rngA.Rows.Count >= ws.Rows.Count) Or (rngB.Rows.Count >= ws.Rows.Count) _
           Or (rngA.Rows.Count > 200000) Or (rngB.Rows.Count > 200000)
    If Not overflow Then Exit Sub

    ' find the last filled cell in column A and take that table
    Set lastRow = ws.Cells(ws.Rows.Count, rngA.Column).End(xlUp)
    If lastRow.Row < 2 Then
        Set lastRow = ws.Cells(ws.Rows.Count, rngB.Column).End(xlUp)
        If lastRow.Row < 2 Then Exit Sub
    End If

    Set blk = SkipNonHeaderRows(lastRow.CurrentRegion)
    If blk Is Nothing Then Exit Sub

    headerRow = blk.Row
    endRow = blk.Row + blk.Rows.Count - 1
    If endRow <= headerRow Then Exit Sub

    Set rngA = ws.Range(ws.Cells(headerRow + 1, rngA.Column), ws.Cells(endRow, rngA.Column))
    Set rngB = ws.Range(ws.Cells(headerRow + 1, rngB.Column), ws.Cells(endRow, rngB.Column))

Done:
End Sub

'--- leave the format-copy waiting state (shared by Esc, timeout and a finished paste) ---
Private Sub ClearFormatBrush()
    On Error Resume Next
    If mFmtSrcOn Then
        Application.OnKey "{ESC}"
        Application.OnTime mFmtSrcTime, MacroPrefix() & "FormatBrushTimeout", , False
    End If
    mFmtSrcOn = False
    mFmtSrcAddr = ""
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub CancelFormatBrush()
    If Not mFmtSrcOn Then
        On Error Resume Next
        Application.OnKey "{ESC}"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    ClearFormatBrush
    Application.StatusBar = "Format brush cancelled."
    ScheduleStatusClear
End Sub

Public Sub FormatBrushTimeout()
    If Not mFmtSrcOn Then Exit Sub
    ClearFormatBrush
    Application.StatusBar = "The format brush " & FORMAT_BRUSH_SECONDS & " seconds passed, so it cleared itself."
    ScheduleStatusClear
End Sub

'--- a column number into column letters (A, B, ... AA) ---
Private Function ColumnLetter(ByVal colNo As Long) As String
    Dim s As String
    Dim n As Long
    n = colNo
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColumnLetter = s
End Function

'------------------------------------------------------------
' reset the Text to Columns settings
'   Excel remembers the last TextToColumns settings and
'   uses them for later [Paste] operations as well.
'   leave every delimiter off and pasting tab-separated values
'   puts everything into one cell. so they are reset to the default (tab).
'------------------------------------------------------------
Private Sub RestoreTextToColumns()
    Dim ws As Worksheet
    Dim tmp As Range

    On Error GoTo Done
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Visible = xlSheetVeryHidden

    Set tmp = ws.Range("A1")
    tmp.Value = "a" & vbTab & "b"
    tmp.TextToColumns Destination:=tmp, _
                       DataType:=xlDelimited, _
                       TextQualifier:=xlDoubleQuote, _
                       ConsecutiveDelimiter:=False, _
                       Tab:=True, Semicolon:=False, Comma:=False, _
                       Space:=False, Other:=False, _
                       TrailingMinusNumbers:=True

Done:
    On Error Resume Next
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Err.Clear
    On Error GoTo 0
End Sub


'============================================================
'  comma format look settings
'    positive    P : N plain · Y with a +
'    positive color PC: K black · B blue · R red
'    negative    N : T triangle up · V triangle down · P brackets · M minus
'    negative color C : R red · K black · B blue
'    zero        Z : D dash - · Z the digit 0 · N blank
'============================================================

Private Function CommaPos() As String
    CommaPos = ReadSetting("cmaPlus", "N")
End Function
Private Function CommaPosColor() As String
    CommaPosColor = ReadSetting("cmaPColor", "K")
End Function
Private Function CommaNeg() As String
    CommaNeg = ReadSetting("cmaNeg", "T")
End Function
Private Function CommaNegColor() As String
    CommaNegColor = ReadSetting("cmaColor", "R")
End Function
Private Function CommaZero() As String
    CommaZero = ReadSetting("cmaZero", "D")
End Function

Private Function ColorCode(ByVal c As String) As String
    Select Case c
        Case "B": ColorCode = "[Blue]"
        Case "R": ColorCode = "[Red]"
        Case Else: ColorCode = ""
    End Select
End Function

Private Function ColorName(ByVal c As String) As String
    Select Case c
        Case "B": ColorName = "Blue"
        Case "R": ColorName = "red"
        Case Else: ColorName = "black"
    End Select
End Function

'--- the actual number format string for the current settings ---
Private Function CommaFormatString() As String
    Dim posFmt As String, negFmt As String, zeroFmt As String

    posFmt = ColorCode(CommaPosColor()) & IIf(CommaPos() = "Y", "+", "") & "#,##0"

    Select Case CommaNeg()
        Case "V": negFmt = ColorCode(CommaNegColor()) & "▽#,##0"
        Case "P": negFmt = ColorCode(CommaNegColor()) & "(#,##0)"
        Case "M": negFmt = ColorCode(CommaNegColor()) & "-#,##0"
        Case Else: negFmt = ColorCode(CommaNegColor()) & "△#,##0"
    End Select

    Select Case CommaZero()
        Case "Z": zeroFmt = "0"
        Case "N": zeroFmt = """"""
        Case Else: zeroFmt = "-"
    End Select

    CommaFormatString = posFmt & ";" & negFmt & ";" & zeroFmt
End Function

'--- a visible example ---
Private Function PosPreview() As String
    PosPreview = IIf(CommaPos() = "Y", "+3,333", "3,333")
End Function

Private Function NegPreview() As String
    Select Case CommaNeg()
        Case "V": NegPreview = "▽3,333"
        Case "P": NegPreview = "(3,333)"
        Case "M": NegPreview = "-3,333"
        Case Else: NegPreview = "△3,333"
    End Select
End Function

Private Function ZeroPreview() As String
    Select Case CommaZero()
        Case "Z": ZeroPreview = "0"
        Case "N": ZeroPreview = "blank"
        Case Else: ZeroPreview = "-"
    End Select
End Function

Private Function CommaFormatOneLine() As String
    CommaFormatOneLine = PosPreview() & "  /  " & NegPreview() & "  /  " & ZeroPreview()
End Function

Private Function NegName(ByVal c As String) As String
    Select Case c
        Case "V": NegName = "inverted triangle ▼"
        Case "P": NegName = "parentheses ( )"
        Case "M": NegName = "minus -"
        Case Else: NegName = "triangle ▲"
    End Select
End Function

Private Function ZeroName(ByVal c As String) As String
    Select Case c
        Case "Z": ZeroName = "zero 0"
        Case "N": ZeroName = "blank"
        Case Else: ZeroName = "dash -"
    End Select
End Function

Private Function NextColor(ByVal c As String) As String
    Select Case c
        Case "K": NextColor = "B"
        Case "B": NextColor = "R"
        Case Else: NextColor = "K"
    End Select
End Function
Private Function NextNeg(ByVal c As String) As String
    Select Case c
        Case "T": NextNeg = "V"
        Case "V": NextNeg = "P"
        Case "P": NextNeg = "M"
        Case Else: NextNeg = "T"
    End Select
End Function
Private Function NextZero(ByVal c As String) As String
    Select Case c
        Case "D": NextZero = "Z"
        Case "Z": NextZero = "N"
        Case Else: NextZero = "D"
    End Select
End Function

'------------------------------------------------------------
' comma format look - submenu
'------------------------------------------------------------
Public Sub CommaFormatMenu()
    Dim cb As CommandBar
    Dim btn As CommandBarButton

    On Error Resume Next
    Application.CommandBars(BAR_COMMA_FORMAT).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=BAR_COMMA_FORMAT, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    AddMenuTitle cb, "Comma format look  (apply with Ctrl+Shift+X)"

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     Preview"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "          Positive  " & PosPreview() & "   (" & ColorName(CommaPosColor()) & ")"
    btn.Style = msoButtonCaption
    btn.Enabled = False

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "          Negative  " & NegPreview() & "   (" & ColorName(CommaNegColor()) & ")"
    btn.Style = msoButtonCaption
    btn.Enabled = False

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "          0        " & ZeroPreview()
    btn.Style = msoButtonCaption
    btn.Enabled = False

    '--- the line that changes it ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  Positive   " & IIf(CommaPos() = "Y", "with +", "plain") & _
                  "   ->   " & IIf(CommaPos() = "Y", "plain", "with +")
    btn.OnAction = MacroPrefix() & "ToggleCommaPos"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  Positive color " & ColorName(CommaPosColor()) & "   ->   " & ColorName(NextColor(CommaPosColor()))
    btn.OnAction = MacroPrefix() & "ToggleCommaPosColor"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  Negative   " & NegName(CommaNeg()) & "   ->   " & NegName(NextNeg(CommaNeg()))
    btn.OnAction = MacroPrefix() & "ToggleCommaNeg"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  Negative color " & ColorName(CommaNegColor()) & "   ->   " & ColorName(NextColor(CommaNegColor()))
    btn.OnAction = MacroPrefix() & "ToggleCommaNegColor"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  0          " & ZeroName(CommaZero()) & "   ->   " & ZeroName(NextZero(CommaZero()))
    btn.OnAction = MacroPrefix() & "ToggleCommaZero"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  Back to defaults   (black 3,333 / red ▲3,333 / -)"
    btn.OnAction = MacroPrefix() & "CommaFormatDefaults"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "Could not show the menu." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Comma Format"
End Sub

Public Sub ToggleCommaPos()
    WriteSetting "cmaPlus", IIf(CommaPos() = "Y", "N", "Y")
    CommaFormatMenu
End Sub
Public Sub ToggleCommaPosColor()
    WriteSetting "cmaPColor", NextColor(CommaPosColor())
    CommaFormatMenu
End Sub
Public Sub ToggleCommaNeg()
    WriteSetting "cmaNeg", NextNeg(CommaNeg())
    CommaFormatMenu
End Sub
Public Sub ToggleCommaNegColor()
    WriteSetting "cmaColor", NextColor(CommaNegColor())
    CommaFormatMenu
End Sub
Public Sub ToggleCommaZero()
    WriteSetting "cmaZero", NextZero(CommaZero())
    CommaFormatMenu
End Sub
Public Sub CommaFormatDefaults()
    WriteSetting "cmaPlus", "N"
    WriteSetting "cmaPColor", "K"
    WriteSetting "cmaNeg", "T"
    WriteSetting "cmaColor", "R"
    WriteSetting "cmaZero", "D"
    CommaFormatMenu
End Sub
