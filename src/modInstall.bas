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

Private Sub EnsureDir(ByVal targetDir As String)
    Dim parentDir As String
    Dim p As Long

    If Len(targetDir) = 0 Then Exit Sub
    If Len(Dir(targetDir, vbDirectory)) > 0 Then Exit Sub

    If Right$(targetDir, 1) = "\" Then targetDir = Left$(targetDir, Len(targetDir) - 1)

    ' 드라이브 루트(C:)까지 올라가면 더 만들지 않는다
    If Len(targetDir) <= 2 Then Exit Sub

    ' UNC(\\server\share) 는 서버/공유 단계까지는 만들지 않는다
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

Public Sub PERSONAL_설치()
    Dim srcPath As String
    Dim dstDir As String
    Dim dstPath As String
    Dim originalAlerts As Boolean
    Dim startTime As Double

    On Error GoTo ErrHandler

    srcPath = GetSrcPath()
    dstDir = GetDstDir()
    dstPath = GetDstPath()

    originalAlerts = Application.DisplayAlerts

    ' 1. 원본 확인
    If Len(Dir(srcPath)) = 0 Then
        MsgBox "이 파일과 같은 폴더에 PERSONAL.XLSB 이 없습니다." & vbCrLf & vbCrLf & _
               "현재 위치: " & ThisWorkbook.Path & vbCrLf & vbCrLf & _
               "두 파일을 같은 폴더에 두고 다시 실행해주세요.", vbExclamation, "설치 실패"
        Exit Sub
    End If

    If FileLen(srcPath) = 0 Then
        MsgBox "PERSONAL.XLSB 원본 파일이 손상된 것으로 보입니다(크기 0)." & vbCrLf & _
               "설치를 중단합니다.", vbExclamation, "설치 실패"
        Exit Sub
    End If

    ' 2. 대상 폴더 준비
    EnsureDir dstDir
    If Len(Dir(dstDir, vbDirectory)) = 0 Then
        MsgBox "대상 폴더를 만들 수 없습니다." & vbCrLf & dstDir, vbExclamation, "설치 실패"
        Exit Sub
    End If

    ' 3. 열려있는 PERSONAL.XLSB 닫기
    ClosePersonalIfOpen

    ' 4. 기존 파일 삭제 (백업 없음)
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
            MsgBox "PERSONAL.XLSB 가 다른 Excel 창에서 사용 중일 수 있습니다." & vbCrLf & vbCrLf & _
                   "Excel 창을 모두 종료한 후 다시 설치해주세요.", vbExclamation, "설치 실패"
            Application.DisplayAlerts = originalAlerts
            Exit Sub
        End If
    End If

    ' 5. 복사
    FileCopy srcPath, dstPath

    ' 6. 검증
    If Len(Dir(dstPath)) = 0 Then
        MsgBox "파일 복사에 실패했습니다.", vbExclamation, "설치 실패"
        Application.DisplayAlerts = originalAlerts
        Exit Sub
    End If

    If FileLen(srcPath) <> FileLen(dstPath) Or FileLen(dstPath) = 0 Then
        On Error Resume Next
        Kill dstPath
        Err.Clear
        On Error GoTo ErrHandler
        MsgBox "복사된 파일 검증에 실패했습니다. 설치를 중단합니다.", vbExclamation, "설치 실패"
        Application.DisplayAlerts = originalAlerts
        Exit Sub
    End If

    ' 7. 성공 메시지
    Application.DisplayAlerts = originalAlerts
    MsgBox "설치가 완료되었습니다." & vbCrLf & vbCrLf & _
           "Excel 창을 모두 종료한 후 다시 실행해주세요." & vbCrLf & _
           "다음 실행부터 업무용 매크로가 적용됩니다." & vbCrLf & vbCrLf & _
           "설치 위치: " & dstPath, vbInformation, "설치 완료"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = originalAlerts
    MsgBox "설치 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & _
           "오류 내용: " & Err.Description, vbCritical, "설치 오류"
End Sub

Public Sub PERSONAL_제거()
    Dim dstPath As String
    Dim ans As VbMsgBoxResult
    Dim startTime As Single

    On Error GoTo ErrHandler

    dstPath = GetDstPath()

    If Len(Dir(dstPath)) = 0 Then
        MsgBox "설치된 PERSONAL.XLSB 이 없습니다.", vbInformation, "제거"
        Exit Sub
    End If

    ans = MsgBox("설치된 PERSONAL.XLSB 를 삭제하시겠습니까?" & vbCrLf & vbCrLf & dstPath, _
                 vbYesNo + vbQuestion, "제거 확인")
    If ans <> vbYes Then Exit Sub

    ' 현재 Excel 에 열려 있으면 먼저 닫는다
    ClosePersonalIfOpen

    On Error Resume Next
    SetAttr dstPath, vbNormal
    Kill dstPath
    Err.Clear
    On Error GoTo ErrHandler

    ' 파일 핸들이 늦게 풀릴 수 있으므로 잠시 기다린다
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
        MsgBox "PERSONAL.XLSB 가 다른 Excel 창에서 사용 중일 수 있습니다." & vbCrLf & vbCrLf & _
               "Excel 창을 모두 종료한 후 다시 삭제해주세요.", vbExclamation, "제거 실패"
        Exit Sub
    End If

    MsgBox "삭제가 완료되었습니다." & vbCrLf & vbCrLf & _
           "Excel 창을 모두 종료한 후 다시 실행하면" & vbCrLf & _
           "업무도구 단축키가 사라집니다.", vbInformation, "제거 완료"
    Exit Sub

ErrHandler:
    MsgBox "제거 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & _
           "오류 내용: " & Err.Description, vbCritical, "제거 오류"
End Sub
