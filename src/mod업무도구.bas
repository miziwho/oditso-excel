Attribute VB_Name = "mod업무도구"
Option Explicit

'============================================================
'  Excel 업무도구  v1.0.0   (2026-08-25)
'  PERSONAL.XLSB 전용 단일 모듈
'
'  설치 : VBE에서 [파일 가져오기]로 이 파일 하나만 넣고,
'         ThisWorkbook_코드.txt 내용을 ThisWorkbook에 붙여넣으면 끝.
'  설정 : 시트 없이 PERSONAL.XLSB의 이름정의에 저장한다.
'
'  단축키
'    Ctrl+Shift+Q  새 창으로 복사
'    Ctrl+Shift+X  쉼표 형식
'    Ctrl+Shift+W  행/열 자동맞춤
'    Ctrl+Shift+S  첫 행 고정 / 해제
'    Ctrl+Shift+Y  텍스트 나누기 - 일반으로
'    Ctrl+Shift+I  텍스트 나누기 - 텍스트로
'    Ctrl+Shift+H  시트 일괄정리
'    Ctrl+Shift+J  피벗테이블 만들기 (새 시트)
'    Ctrl+Shift+T  클래식 피벗 정리
'    Ctrl+Shift+D  중복행 색칠
'    Ctrl+Shift+C  셀 배경 지우기
'    Ctrl+Shift+U  표 테두리
'    Ctrl+Shift+A  필터 (자동필터 토글 / 고급필터 해제)
'    Ctrl+Shift+F  고급필터 실행
'    Ctrl+Shift+R  조건 열 세팅
'    Ctrl+Shift+R  필터 해제 (자동/고급 공용)
'    Ctrl+Shift+E  시트 빼서 텍스트 저장
'    Ctrl+Shift+N  파일 경로 복사
'    Ctrl+Shift+Z  선택한 셀 값으로 필터
'    Ctrl+Shift+P  암호 설정
'    Ctrl+Shift+O  새 창 열기
'    Ctrl+Shift+K  단축키 목록 붙이기 / 지우기
'    Ctrl+Shift+M  업무도구 메뉴
'    (자동)        저장된 파일을 닫을 때 열기암호 확인
'============================================================

Public Const 도구버전 As String = "1.0.0"
Public Const 도구버전일자 As String = "2026-08-25"
Public Const 제작자 As String = "SH"

Private Const 기본암호_초기값 As String = "1234"
Private Const 암호최대길이 As Long = 15
Private Const 메뉴이름 As String = "업무도구메뉴"
Private Const 치트시트이름 As String = "업무도구_단축키목록"

Private mBusy As Boolean               ' 재진입 방지
Private mSuppressOnce As Boolean       ' 지연 닫기 중 재질문 방지
Private mPendingCloseName As String    ' 지연 닫기 대상

' SUBTOTAL 합계 되돌리기용 보관
Private mUndoWb As String
Private mUndoWs As String
Private mUndoAddr As String
Private mUndoFormula As Variant
Private mUndoBold As Variant

' 중복행 색칠 되돌리기용 보관
Private mDupWb As String
Private mDupWs As String
Private mDupAddr As String
Private mDupRows As Variant
Private mDupColor As Variant
Private mDupIndex As Variant
Private mDupCount As Long

' 단축키 목록 되돌리기용 보관
Private mCheatWb As String
Private mCheatWs As String

' 텍스트 나누기 되돌리기용 보관
Private mTxtWb As String
Private mTxtWs As String
Private mTxtAddr As String
Private mTxtValue As Variant
Private mTxtFormat As Variant
Private mTxtOn As Boolean

' 쉼표 형식 / 행열 자동맞춤 / 조건 열 세팅 되돌리기용 보관
Private mFmtWb As String, mFmtWs As String, mFmtAddr As String
Private mFmtValue As Variant, mFmtOn As Boolean

Private mFitWb As String, mFitWs As String, mFitAddr As String
Private mFitColW As Variant, mFitRowH As Variant, mFitOn As Boolean

Private mCondWb As String, mCondWs As String, mCondAddr As String
Private mCondF As Variant, mCondN As Variant
Private mCondC As Variant, mCondI As Variant
Private mCondB As Variant, mCondFC As Variant
Private mCondOn As Boolean

' 표 테두리 되돌리기용 보관
Private mBdrWb As String, mBdrWs As String, mBdrAddr As String
Private mBdrLine As Variant, mBdrBold As Variant, mBdrAlign As Variant
Private mBdrOn As Boolean

' 셀 배경 지우기 되돌리기용 보관
Private mBgWb As String, mBgWs As String, mBgAddr As String
Private mBgColor As Variant, mBgIndex As Variant, mBgOn As Boolean

' SUBTOTAL 토글 위치
Private mUndoOn As Boolean


'============================================================
'  A. 설정 저장 / 읽기  (이름정의 사용 - 시트 불필요)
'============================================================

Private Function 설정읽기(ByVal 키 As String, ByVal 기본값 As String) As String
    Dim s As String

    On Error Resume Next
    s = CStr(ThisWorkbook.Names("cfg_" & 키).RefersTo)
    Err.Clear
    On Error GoTo 0

    ' RefersTo 는  ="값"  형태로 저장된다
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

    If Len(s) = 0 Then s = 기본값
    설정읽기 = s
End Function

Private Sub 설정쓰기(ByVal 키 As String, ByVal 값 As String)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="cfg_" & 키, _
                           RefersTo:="=""" & Replace(값, """", """""") & """", _
                           Visible:=False
    ThisWorkbook.Save
    Err.Clear
    On Error GoTo 0
End Sub

Public Function 기본암호읽기() As String
    기본암호읽기 = 설정읽기("pwd", 기본암호_초기값)
End Function

Public Function 자동확인사용() As Boolean
    자동확인사용 = (설정읽기("auto", "1") <> "0")
End Function


'============================================================
'  B. 초기화 / 단축키
'============================================================

Private Function 접두어() As String
    접두어 = "'" & ThisWorkbook.Name & "'!"
End Function

Public Sub 단축키등록()
    Dim p As String
    p = 접두어()

    On Error Resume Next
    Application.OnKey "^+q", p & "새창으로복사"
    Application.OnKey "^+x", p & "쉼표형식"
    Application.OnKey "^+w", p & "행열자동맞춤"
    Application.OnKey "^+s", p & "첫행고정"
    Application.OnKey "^+y", p & "일반으로변환"
    Application.OnKey "^+i", p & "텍스트로변환"
    Application.OnKey "^+h", p & "시트일괄정리"
    Application.OnKey "^+j", p & "피벗만들기"
    Application.OnKey "^+t", p & "클래식피벗"
    Application.OnKey "^+d", p & "중복행색칠"
    Application.OnKey "^+c", p & "중복색지우기"
    Application.OnKey "^+u", p & "표테두리"
    Application.OnKey "^+a", p & "자동필터"
    Application.OnKey "^+f", p & "고급필터"
    Application.OnKey "^+r", p & "조건블록만들기"
    Application.OnKey "^+g", p & "소계넣기"
    Application.OnKey "^+e", p & "시트빼서텍스트저장"
    Application.OnKey "^+n", p & "파일경로복사"
    Application.OnKey "^+z", p & "선택값으로필터"
    Application.OnKey "^+p", p & "암호설정"
    Application.OnKey "^+o", p & "새창열기"
    Application.OnKey "^+k", p & "치트시트"
    Application.OnKey "^+m", p & "매크로런처"
    Err.Clear
    On Error GoTo 0

    mBusy = False
    mSuppressOnce = False
    mPendingCloseName = ""
End Sub

Public Sub 단축키해제()
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
    Application.OnKey "^+m"
    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------
' 도구가 안 먹을 때 수동 재초기화 (Alt+F8 에서 실행)
'------------------------------------------------------------
Public Sub 업무도구_다시시작()
    On Error Resume Next
    ThisWorkbook.이벤트연결
    Err.Clear
    On Error GoTo 0
    단축키등록
    MsgBox "업무도구를 다시 초기화했습니다." & vbCrLf & _
           "버전 " & 도구버전, vbInformation, "업무도구"
End Sub

'------------------------------------------------------------
' 진단 (Alt+F8 에서 실행) - 문제 생겼을 때 상태 확인용
'------------------------------------------------------------
Public Sub 업무도구_진단()
    Dim s As String
    Dim 이벤트 As String

    이벤트 = "확인 불가"
    On Error Resume Next
    이벤트 = IIf(ThisWorkbook.이벤트연결됨(), "연결됨", "끊김")
    Err.Clear
    On Error GoTo 0

    s = "Excel 업무도구 " & 도구버전 & "  (" & 도구버전일자 & ")" & vbCrLf & _
        "made by " & 제작자 & vbCrLf & vbCrLf & _
        "Excel 버전      : " & Application.Version & _
                              " (" & IIf(사용비트() = 64, "64bit", "32bit") & ")" & vbCrLf & _
        "설치 위치       : " & ThisWorkbook.FullName & vbCrLf & _
        "닫기 이벤트     : " & 이벤트 & vbCrLf & _
        "자동 암호확인   : " & IIf(자동확인사용(), "켜짐", "꺼짐") & vbCrLf & _
        "기본 암호       : " & 기본암호읽기() & vbCrLf & vbCrLf & _
        "단축키 미작동 시 [업무도구_다시시작]을 실행해보세요."

    MsgBox s, vbInformation, "업무도구 진단"
End Sub

Private Function 사용비트() As Long
    #If Win64 Then
        사용비트 = 64
    #Else
        사용비트 = 32
    #End If
End Function


'============================================================
'  C. 열기 암호 엔진  (수동 / 자동 공용)
'============================================================

'------------------------------------------------------------
' 암호 입력받기   반환 0 = 정상 / 1 = 취소
'------------------------------------------------------------
Private Function 암호입력받기(ByRef Pwd As String) As Long
    Dim s1 As String, s2 As String
    Dim def As String
    Dim r As VbMsgBoxResult

    Pwd = ""
    def = 기본암호읽기()

    Do
        s1 = InputBox("설정할 열기 암호를 입력하세요. (1~" & 암호최대길이 & "자)", _
                      "열기 암호 설정", def)

        If StrPtr(s1) = 0 Then                          ' [취소] 버튼
            암호입력받기 = 1
            Exit Function
        End If

        If Len(s1) = 0 Then
            r = MsgBox("빈 암호는 사용할 수 없습니다." & vbCrLf & _
                       "다시 입력하시겠습니까?", vbOKCancel + vbExclamation, "열기 암호 설정")
            If r = vbCancel Then
                암호입력받기 = 1
                Exit Function
            End If
        ElseIf Len(s1) > 암호최대길이 Then
            MsgBox "암호는 " & 암호최대길이 & "자 이하로 입력해주세요.", _
                   vbExclamation, "열기 암호 설정"
        Else
            Exit Do
        End If
    Loop

    ' 기본 암호를 그대로 쓴 경우에는 재입력 확인을 생략
    If StrComp(s1, def, vbBinaryCompare) = 0 Then
        Pwd = s1
        암호입력받기 = 0
        Exit Function
    End If

    Do
        s2 = InputBox("같은 암호를 한 번 더 입력하세요.", "암호 확인")

        If StrPtr(s2) = 0 Then
            암호입력받기 = 1
            Exit Function
        End If

        If StrComp(s1, s2, vbBinaryCompare) = 0 Then Exit Do

        r = MsgBox("두 암호가 일치하지 않습니다. (대소문자를 구분합니다.)" & vbCrLf & _
                   "다시 입력하시겠습니까?", vbOKCancel + vbExclamation, "암호 확인")
        If r = vbCancel Then
            암호입력받기 = 1
            Exit Function
        End If
    Loop

    Pwd = s1
    암호입력받기 = 0
End Function

'------------------------------------------------------------
' 열기 암호를 실제 파일에 기록
'   같은 경로 / 같은 파일형식으로 SaveAs 하고 결과를 검증한다.
'------------------------------------------------------------
Private Function 열기암호적용(ByVal Wb As Workbook, _
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
        ErrMsg = "대상 통합문서를 찾을 수 없습니다."
        GoTo Restore
    End If
    If Wb.Path = "" Then
        ErrMsg = "아직 저장된 적이 없는 파일입니다."
        GoTo Restore
    End If

    sFull = Wb.FullName
    lFmt = Wb.FileFormat

    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Wb.SaveAs Filename:=sFull, FileFormat:=lFmt, Password:=Pwd, CreateBackup:=False

    '--- 저장 결과 검증 ---
    ok = Wb.HasPassword

    fileOK = True
    On Error Resume Next
    fileOK = (Len(Dir$(sFull)) > 0)
    If Err.Number <> 0 Then
        fileOK = True                       ' Dir 자체가 실패하면 판정 보류
        Err.Clear
    End If
    On Error GoTo Fail

    If Not fileOK Then ok = False
    If Not ok Then ErrMsg = "저장은 실행되었으나 열기 암호 적용이 확인되지 않았습니다."

Restore:
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    열기암호적용 = ok
    Exit Function

Fail:
    ErrMsg = "오류 번호: " & Err.Number & vbCrLf & "내용: " & Err.Description
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    열기암호적용 = False
End Function

'------------------------------------------------------------
' 암호설정 (Ctrl+Shift+X) - 수동 실행
'------------------------------------------------------------
Public Sub 암호설정()
    Dim Wb As Workbook
    Dim Pwd As String
    Dim ErrMsg As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "암호를 걸 파일을 먼저 선택한 뒤 실행하세요.", vbExclamation, "열기 암호 설정"
        Exit Sub
    End If

    Set Wb = ActiveWorkbook

    If StrComp(Wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then
        MsgBox "암호를 걸 파일을 먼저 선택한 뒤 실행하세요.", vbExclamation, "열기 암호 설정"
        Exit Sub
    End If
    If Wb.Path = "" Then
        MsgBox "아직 저장되지 않은 파일입니다." & vbCrLf & _
               "먼저 파일을 저장한 뒤 다시 실행해주세요.", vbExclamation, "열기 암호 설정"
        Exit Sub
    End If
    If Wb.ReadOnly Then
        MsgBox "읽기 전용으로 열린 파일에는 암호를 설정할 수 없습니다.", _
               vbExclamation, "열기 암호 설정"
        Exit Sub
    End If
    If Wb.HasPassword Then
        MsgBox "이 파일에는 이미 열기 암호가 설정되어 있습니다.", _
               vbInformation, "열기 암호 설정"
        Exit Sub
    End If

    If 암호입력받기(Pwd) <> 0 Then Exit Sub

    If 열기암호적용(Wb, Pwd, ErrMsg) Then
        MsgBox "열기 암호를 설정하고 저장했습니다." & vbCrLf & vbCrLf & Wb.Name, _
               vbInformation, "열기 암호 설정"
    Else
        MsgBox "암호 설정 또는 저장 중 오류가 발생했습니다." & vbCrLf & vbCrLf & ErrMsg, _
               vbCritical, "열기 암호 설정"
    End If
End Sub


'============================================================
'  D. 닫을 때 자동 암호 확인
'============================================================

Public Sub 자동확인전환()
    Dim 현재 As Boolean
    현재 = 자동확인사용()
    설정쓰기 "auto", IIf(현재, "0", "1")

    If 현재 Then
        MsgBox "닫을 때 암호 확인을 껐습니다." & vbCrLf & _
               "같은 메뉴에서 다시 켤 수 있습니다.", vbInformation, "업무도구"
    Else
        MsgBox "저장된 파일을 닫을 때 암호 설정 여부를 다시 묻습니다.", _
               vbInformation, "업무도구"
    End If
End Sub

Public Sub 기본암호변경()
    Dim s As String

    s = InputBox("암호 입력창에 미리 채워 넣을 기본 암호입니다. (1~" & 암호최대길이 & "자)", _
                 "기본 암호 변경", 기본암호읽기())

    If StrPtr(s) = 0 Then Exit Sub
    If Len(s) = 0 Then
        MsgBox "빈 값은 사용할 수 없습니다.", vbExclamation, "기본 암호 변경"
        Exit Sub
    End If
    If Len(s) > 암호최대길이 Then
        MsgBox "암호는 " & 암호최대길이 & "자 이하로 입력해주세요.", vbExclamation, "기본 암호 변경"
        Exit Sub
    End If

    설정쓰기 "pwd", s
    MsgBox "기본 암호를 '" & s & "' 로 변경했습니다.", vbInformation, "기본 암호 변경"
End Sub

'------------------------------------------------------------
' 검사 대상 통합문서인지 판정
'------------------------------------------------------------
Private Function 대상통합문서인가(ByVal Wb As Workbook) As Boolean
    Dim ext As String
    Dim p As Long

    On Error GoTo Fail

    If Wb Is Nothing Then Exit Function
    If StrComp(Wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then Exit Function
    If Wb.Path = "" Then Exit Function                  ' 저장된 적 없는 새 문서
    If Wb.IsAddin Then Exit Function
    If Wb.ReadOnly Then Exit Function

    p = InStrRev(Wb.Name, ".")
    If p = 0 Then Exit Function
    ext = LCase$(Mid$(Wb.Name, p + 1))

    Select Case ext
        Case "xlsx", "xlsm", "xlsb", "xls"
            대상통합문서인가 = True
    End Select

    Exit Function
Fail:
    대상통합문서인가 = False
End Function

'------------------------------------------------------------
' 닫기 직전 처리  (ThisWorkbook 의 App_WorkbookBeforeClose 에서 호출)
'------------------------------------------------------------
Public Sub 닫기전_암호확인(ByVal Wb As Workbook, ByRef Cancel As Boolean)
    Dim ans As VbMsgBoxResult
    Dim Pwd As String
    Dim ErrMsg As String
    Dim msg As String

    On Error GoTo SafeExit

    If mBusy Then Exit Sub
    If mSuppressOnce Then Exit Sub
    If Not 자동확인사용() Then Exit Sub
    If Not 대상통합문서인가(Wb) Then Exit Sub
    If Wb.HasPassword Then Exit Sub

    msg = "현재 문서에 열기 암호가 설정되어 있지 않습니다." & vbCrLf & vbCrLf & _
          "  " & Wb.Name & vbCrLf & vbCrLf & _
          "[예]      암호를 설정하고 닫습니다." & vbCrLf & _
          "[아니오]  그냥 닫습니다." & vbCrLf & _
          "[취소]    닫지 않고 그대로 둡니다.  (Esc)"

    ans = MsgBox(msg, vbYesNoCancel + vbQuestion, "열기 암호 확인")

    If ans = vbNo Then Exit Sub                         ' 원래 닫기 동작 그대로

    If ans = vbCancel Then                              ' 닫기 자체를 취소
        Cancel = True
        Exit Sub
    End If

    '--- [예] : 닫기를 잠시 멈추고 암호 설정 ---
    Cancel = True
    mBusy = True

    If 암호입력받기(Pwd) <> 0 Then
        mBusy = False
        Exit Sub                                        ' 파일을 열린 상태로 유지
    End If

    If Not 열기암호적용(Wb, Pwd, ErrMsg) Then
        mBusy = False
        MsgBox "암호 설정 또는 저장 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
               ErrMsg & vbCrLf & vbCrLf & _
               "파일을 닫지 않았습니다. 내용을 확인한 뒤 다시 시도해주세요.", _
               vbCritical, "열기 암호 확인"
        Exit Sub                                        ' 실패 시 절대 닫지 않는다
    End If

    MsgBox "열기 암호를 설정하고 저장했습니다.", vbInformation, "열기 암호 확인"

    mPendingCloseName = Wb.Name
    mBusy = False

    ' 이벤트가 끝난 뒤 안전하게 닫는다
    On Error Resume Next
    Application.OnTime Now, 접두어() & "지연닫기실행"
    Err.Clear
    On Error GoTo 0
    Exit Sub

SafeExit:
    mBusy = False
    Cancel = True
    MsgBox "암호 확인 처리 중 오류가 발생했습니다." & vbCrLf & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & _
           "내용: " & Err.Description & vbCrLf & vbCrLf & _
           "파일을 닫지 않았습니다.", vbCritical, "열기 암호 확인"
End Sub

'------------------------------------------------------------
' 지연 닫기 (Application.OnTime 콜백)
'------------------------------------------------------------
Public Sub 지연닫기실행()
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
' 상태표시줄 정리 (Application.OnTime 콜백)
'------------------------------------------------------------
Public Sub 상태표시줄정리()
    On Error Resume Next
    Application.StatusBar = False
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub 상태표시줄예약()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 4), 접두어() & "상태표시줄정리"
    Err.Clear
    On Error GoTo 0
End Sub


'============================================================
'  E. 업무도구
'============================================================

'------------------------------------------------------------
' 1. 새창으로복사 (Ctrl+Shift+Q)
'------------------------------------------------------------
Public Sub 새창으로복사()
    Dim sh As Object

    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then
        MsgBox "복사할 시트가 없습니다.", vbExclamation, "새 창으로 복사"
        Exit Sub
    End If

    Set sh = ActiveSheet
    If sh Is Nothing Then Exit Sub

    sh.Copy
    Exit Sub

Fail:
    MsgBox "새 창으로 복사하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "새 창으로 복사"
End Sub

'------------------------------------------------------------
' 2. 쉼표형식 (Ctrl+Shift+Z)
'------------------------------------------------------------
Public Sub 쉼표형식()
    On Error GoTo Fail

    If TypeName(Selection) <> "Range" Then
        MsgBox "셀을 선택한 뒤 실행해주세요.", vbExclamation, "쉼표 형식"
        Exit Sub
    End If

    ' 되돌리기용 보관
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

    Selection.NumberFormat = "#,##0;[Red]△#,##0;-"

    If mFmtOn Then
        On Error Resume Next
        Application.OnUndo "쉼표 형식 취소", 접두어() & "쉼표형식되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "쉼표 형식을 적용했습니다." & _
                            IIf(mFmtOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    MsgBox "서식을 적용하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "쉼표 형식"
End Sub

'------------------------------------------------------------
' 3. 클래식피벗 (Ctrl+Shift+A)
'------------------------------------------------------------
Public Sub 클래식피벗()
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
        MsgBox "피벗테이블 안의 셀을 선택한 뒤 실행해주세요.", vbExclamation, "클래식 피벗"
        Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False

    With pt
        .InGridDropZones = True                 ' 클래식 피벗테이블 레이아웃
        .RowAxisLayout xlTabularRow             ' 행 필드를 각각 별도 열로
    End With

    For Each pf In pt.RowFields                 ' 모든 행 필드 부분합 제거
        For i = 1 To 12
            On Error Resume Next
            pf.Subtotals(i) = False
            Err.Clear
            On Error GoTo Fail
        Next i
    Next pf

    pt.RepeatAllLabels xlRepeatLabels           ' 모든 항목 레이블 반복

    Application.ScreenUpdating = oldScreen

    MsgBox "피벗 설정 완료!" & vbCrLf & _
           "· 클래식 피벗 표시" & vbCrLf & _
           "· 행 필드 열 분리" & vbCrLf & _
           "· 부분합 전부 제외" & vbCrLf & _
           "· 모든 항목 레이블 반복" & vbCrLf & vbCrLf & _
           "※ Ctrl+Z 로 되돌릴 수 없습니다.", vbInformation, "클래식 피벗"
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "설정 중 오류가 발생했습니다." & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & Err.Description, _
           vbExclamation, "클래식 피벗"
End Sub

'------------------------------------------------------------
' 4. 시트일괄정리 (Ctrl+Shift+H)
'------------------------------------------------------------
Public Sub 시트일괄정리()
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

    Application.StatusBar = "시트 " & n & "개를 정리했습니다."
    상태표시줄예약
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    MsgBox "시트를 정리하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "시트 일괄정리"
End Sub

'------------------------------------------------------------
' 5. 행열자동맞춤 (Ctrl+Shift+W)
'------------------------------------------------------------
Public Sub 행열자동맞춤()
    Dim rng As Range
    Dim oldScreen As Boolean
    Dim i As Long
    Dim 이전너비() As Variant
    Dim 이전높이() As Variant

    If TypeName(Selection) <> "Range" Then
        MsgBox "셀을 선택한 뒤 실행해주세요.", vbExclamation, "행열 자동맞춤"
        Exit Sub
    End If

    If Selection.Cells.CountLarge >= 2 Then
        Set rng = Selection
    Else
        Set rng = ActiveCell.CurrentRegion       ' 합계 행 너비도 함께 맞춘다
    End If

    oldScreen = Application.ScreenUpdating

    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' 되돌리기용 보관
    On Error Resume Next
    ReDim 이전너비(1 To rng.Columns.Count)
    ReDim 이전높이(1 To rng.Rows.Count)
    For i = 1 To rng.Columns.Count
        이전너비(i) = rng.Columns(i).ColumnWidth
    Next i
    For i = 1 To rng.Rows.Count
        이전높이(i) = rng.Rows(i).RowHeight
    Next i
    mFitWb = rng.Worksheet.Parent.Name
    mFitWs = rng.Worksheet.Name
    mFitAddr = rng.Address
    mFitColW = 이전너비
    mFitRowH = 이전높이
    mFitOn = True
    Err.Clear
    On Error GoTo Fail

    rng.Columns.AutoFit
    rng.Rows.AutoFit

    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    Application.OnUndo "행/열 자동맞춤 취소", 접두어() & "자동맞춤되돌리기"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "행/열 너비를 맞췄습니다. (" & rng.Address(False, False) & ")" & _
                            "  |  Ctrl+Z 로 되돌리기"
    상태표시줄예약
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "자동맞춤에 실패했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "행열 자동맞춤"
End Sub

'------------------------------------------------------------
' 6. 시트빼서텍스트저장 (Ctrl+Shift+E)
'    현재 시트를 새 창으로 뺀 뒤
'    [원본파일명_yyyymmdd.txt] / 텍스트(탭으로 분리) 로 저장창을 띄운다.
'------------------------------------------------------------
Public Sub 시트빼서텍스트저장()
    Dim srcWb As Workbook
    Dim sh As Object
    Dim newWb As Workbook
    Dim sBase As String, sDefault As String
    Dim vPath As Variant
    Dim oldAlerts As Boolean
    Dim 단계 As String
    Dim errNo As Long, errDesc As String
    Dim 저장오류 As Long, 저장설명 As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "저장할 시트가 없습니다.", vbExclamation, "시트 빼서 텍스트 저장"
        Exit Sub
    End If

    oldAlerts = Application.DisplayAlerts       ' 오류 시 복원용. 반드시 먼저 확보

    Set srcWb = ActiveWorkbook
    Set sh = ActiveSheet

    ' 복사하기 전에 원본 이름을 확보 (복사 후에는 ActiveWorkbook 이 바뀐다)
    If srcWb.Path = "" Then
        sBase = sh.Name
    Else
        sBase = srcWb.Name
        If InStrRev(sBase, ".") > 1 Then sBase = Left$(sBase, InStrRev(sBase, ".") - 1)
    End If
    sDefault = 파일명정리(sBase) & "_" & Format$(Date, "yyyymmdd") & ".txt"

    On Error GoTo Fail

    단계 = "시트를 새 창으로 복사"
    sh.Copy
    Set newWb = ActiveWorkbook
    If newWb Is Nothing Then
        MsgBox "새 창을 만들지 못했습니다.", vbExclamation, "시트 빼서 텍스트 저장"
        Exit Sub
    End If

    단계 = "저장 위치 선택"
    vPath = Application.GetSaveAsFilename( _
                InitialFileName:=sDefault, _
                FileFilter:="텍스트 (탭으로 분리) (*.txt),*.txt", _
                Title:="다른 이름으로 저장")

    If VarType(vPath) = vbBoolean Then
        Exit Sub                                ' 취소 - 복사된 새 창은 그대로 둔다
    End If

    단계 = "파일 저장"
    Application.DisplayAlerts = False

    ' 1차 : 텍스트(탭으로 분리) - Windows
    On Error Resume Next
    newWb.SaveAs Filename:=CStr(vPath), FileFormat:=xlTextWindows
    저장오류 = Err.Number
    Err.Clear
    On Error GoTo Fail

    ' 2차 : 현재 플랫폼 텍스트로 재시도
    If 저장오류 <> 0 Then
        On Error Resume Next
        newWb.SaveAs Filename:=CStr(vPath), FileFormat:=xlCurrentPlatformText
        저장오류 = Err.Number
        저장설명 = Err.Description
        Err.Clear
        On Error GoTo Fail
    End If

    Application.DisplayAlerts = oldAlerts

    If 저장오류 <> 0 Then
        MsgBox "저장하지 못했습니다." & vbCrLf & vbCrLf & _
               "저장 위치: " & CStr(vPath) & vbCrLf & _
               "오류 " & 저장오류 & ": " & 저장설명 & vbCrLf & vbCrLf & _
               "복사된 새 창은 열어 두었습니다." & vbCrLf & _
               "F12 를 눌러 직접 저장해보세요.", vbExclamation, "시트 빼서 텍스트 저장"
        Exit Sub
    End If

    MsgBox "저장했습니다." & vbCrLf & vbCrLf & newWb.FullName, _
           vbInformation, "시트 빼서 텍스트 저장"
    Exit Sub

Fail:
    errNo = Err.Number                          ' Err.Clear 전에 반드시 확보
    errDesc = Err.Description

    On Error Resume Next
    Application.DisplayAlerts = oldAlerts
    Err.Clear
    On Error GoTo 0

    MsgBox "저장하지 못했습니다." & vbCrLf & vbCrLf & _
           "단계: " & 단계 & vbCrLf & _
           "오류 " & errNo & ": " & errDesc & vbCrLf & _
           IIf(Len(CStr(vPath)) > 0 And VarType(vPath) <> vbBoolean, _
               "저장 위치: " & CStr(vPath) & vbCrLf, "") & vbCrLf & _
           "복사된 새 창은 닫지 않았습니다. 직접 저장해보실 수 있습니다.", _
           vbExclamation, "시트 빼서 텍스트 저장"
End Sub

Private Function 파일명정리(ByVal s As String) As String
    Dim bad As Variant
    Dim i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next i
    s = Trim$(s)
    If Len(s) = 0 Then s = "sheet"
    파일명정리 = s
End Function

'------------------------------------------------------------
' 7. 중복행색칠 (Ctrl+Shift+S)
'------------------------------------------------------------
Public Sub 중복행색칠()
    Dim rng As Range
    Dim dic As Object
    Dim arr As Variant
    Dim key As String
    Dim r As Long, c As Long
    Dim rowCnt As Long, colCnt As Long
    Dim dupRows As Long
    Dim oldScreen As Boolean
    Dim 그룹색 As Object
    Dim 그룹수 As Long
    Dim 팔레트 As Variant
    Dim 대상행() As Variant
    Dim 이전색() As Variant
    Dim 이전색인덱스() As Variant
    Dim 보임() As Boolean
    Dim 보임범위 As Range
    Dim 영역 As Range
    Dim 시작 As Long
    Dim 보이는수 As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "셀을 선택한 뒤 실행해주세요.", vbExclamation, "중복행 색칠"
        Exit Sub
    End If

    If Selection.Cells.CountLarge >= 2 Then
        Set rng = Selection
    Else
        Set rng = 소계행제외(ActiveCell.CurrentRegion)
    End If

    ' 전체 행/열 선택 방어
    On Error Resume Next
    Set rng = Application.Intersect(rng, rng.Worksheet.UsedRange)
    Err.Clear
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "검사할 데이터가 없습니다.", vbExclamation, "중복행 색칠"
        Exit Sub
    End If
    If rng.Areas.Count > 1 Then
        MsgBox "떨어져 있는 여러 범위는 한 번에 검사할 수 없습니다." & vbCrLf & _
               "연속된 한 범위만 선택한 뒤 다시 실행해주세요.", vbExclamation, "중복행 색칠"
        Exit Sub
    End If

    rowCnt = rng.Rows.Count
    colCnt = rng.Columns.Count

    If rowCnt < 2 Then
        MsgBox "2행 이상을 선택해주세요.", vbExclamation, "중복행 색칠"
        Exit Sub
    End If
    If rowCnt > 100000 Then
        If MsgBox(Format$(rowCnt, "#,##0") & "행을 검사합니다. 시간이 걸릴 수 있습니다." & vbCrLf & _
                  "계속할까요?", vbOKCancel + vbQuestion, "중복행 색칠") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating

    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = rng.Value

    '--- 보이는 행만 대상으로 (필터 / 숨긴 행 제외) ---
    ReDim 보임(1 To rowCnt)
    보이는수 = 0
    On Error Resume Next
    Set 보임범위 = rng.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo Fail

    If 보임범위 Is Nothing Then
        For r = 1 To rowCnt
            보임(r) = True
        Next r
        보이는수 = rowCnt
    Else
        For Each 영역 In 보임범위.Areas
            시작 = 영역.Row - rng.Row + 1
            For r = 시작 To 시작 + 영역.Rows.Count - 1
                If r >= 1 And r <= rowCnt Then
                    If Not 보임(r) Then
                        보임(r) = True
                        보이는수 = 보이는수 + 1
                    End If
                End If
            Next r
        Next 영역
    End If

    If 보이는수 < 2 Then
        Application.ScreenUpdating = oldScreen
        MsgBox "보이는 행이 2행 미만입니다." & vbCrLf & _
               "필터를 확인한 뒤 다시 실행해주세요.", vbExclamation, "중복행 색칠"
        Exit Sub
    End If

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set 그룹색 = CreateObject("Scripting.Dictionary")
    그룹색.CompareMode = 1

    팔레트 = Array(RGB(255, 235, 156), RGB(198, 239, 206), RGB(189, 215, 238), _
                   RGB(255, 199, 206), RGB(226, 214, 244), RGB(255, 224, 178), _
                   RGB(200, 232, 232), RGB(233, 218, 200), RGB(212, 230, 180), _
                   RGB(244, 204, 228))

    ' 1차 : 키별 건수 집계 (보이는 행만)
    For r = 1 To rowCnt
      If 보임(r) Then
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

    ' 2차 : 중복 행 번호 수집 (보이는 행만)
    ReDim 대상행(1 To rowCnt)
    For r = 1 To rowCnt
      If 보임(r) Then
        key = ""
        For c = 1 To colCnt
            key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            dupRows = dupRows + 1
            대상행(dupRows) = r
        End If
      End If
    Next r

    If dupRows > 0 Then
        ' 되돌리기용으로 기존 색 보관
        ReDim 이전색(1 To dupRows, 1 To colCnt)
        ReDim 이전색인덱스(1 To dupRows, 1 To colCnt)
        For r = 1 To dupRows
            For c = 1 To colCnt
                이전색(r, c) = rng.Cells(대상행(r), c).Interior.Color
                이전색인덱스(r, c) = rng.Cells(대상행(r), c).Interior.ColorIndex
            Next c
        Next r
        mDupWb = rng.Worksheet.Parent.Name
        mDupWs = rng.Worksheet.Name
        mDupAddr = rng.Address
        mDupRows = 대상행
        mDupColor = 이전색
        mDupIndex = 이전색인덱스
        mDupCount = dupRows
    End If

    ' 3차 : 그룹별 색 입히기 (보이는 행만)
    dupRows = 0
    For r = 1 To rowCnt
      If 보임(r) Then
        key = ""
        For c = 1 To colCnt
            key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            If Not 그룹색.Exists(key) Then
                그룹색.Add key, 팔레트(그룹수 Mod (UBound(팔레트) + 1))
                그룹수 = 그룹수 + 1
            End If
            rng.Rows(r).Interior.Color = 그룹색(key)
            dupRows = dupRows + 1
        End If
      End If
    Next r

    Application.ScreenUpdating = oldScreen

    If dupRows > 0 Then
        On Error Resume Next
        Application.OnUndo "중복행 색칠 취소", 접두어() & "중복색되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    If dupRows = 0 Then
        MsgBox "중복된 행이 없습니다." & vbCrLf & _
               "검사 범위: " & rng.Address(False, False) & _
               IIf(보이는수 < rowCnt, "  (보이는 " & 보이는수 & "행만 검사)", ""), _
               vbInformation, "중복행 색칠"
    Else
        MsgBox "중복 " & 그룹수 & "종류 / " & Format$(dupRows, "#,##0") & "행을 표시했습니다." & vbCrLf & _
               "같은 내용끼리 같은 색으로 묶었습니다." & vbCrLf & _
               "검사 범위: " & rng.Address(False, False) & _
               IIf(보이는수 < rowCnt, "  (보이는 " & 보이는수 & "행만 검사)", "") & _
               IIf(그룹수 > UBound(팔레트) + 1, vbCrLf & "종류가 " & (UBound(팔레트) + 1) & _
                   "가지를 넘어 일부 색이 반복됩니다.", "") & vbCrLf & vbCrLf & _
               "Ctrl+Z 로 되돌릴 수 있습니다." & vbCrLf & _
               "색만 지우려면 : 범위 선택 후 Ctrl+Shift+C", _
               vbInformation, "중복행 색칠"
    End If
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "중복 검사 중 오류가 발생했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "중복행 색칠"
End Sub

Public Sub 중복색지우기()
    Dim rng As Range
    Dim 보임범위 As Range
    Dim 대상 As Range
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "배경색을 지울 범위를 선택한 뒤 실행해주세요.", vbExclamation, "셀 배경 지우기"
        Exit Sub
    End If

    Set rng = Selection

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    '--- 보이는 셀만 대상 (필터 / 숨긴 행 제외) ---
    On Error Resume Next
    Set 보임범위 = rng.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo Fail

    If 보임범위 Is Nothing Then
        Set 대상 = rng
    Else
        Set 대상 = 보임범위
    End If

    ' 되돌리기용 보관
    배경색보관 대상

    대상.Interior.ColorIndex = xlColorIndexNone

    Application.ScreenUpdating = oldScreen

    If mBgOn Then
        On Error Resume Next
        Application.OnUndo "셀 배경 지우기 취소", 접두어() & "배경색되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "셀 배경색을 지웠습니다. (" & 대상.Address(False, False) & ")" & _
                            IIf(mBgOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "배경색을 지우지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "셀 배경 지우기"
End Sub

'------------------------------------------------------------
' 8. 파일경로복사 (Ctrl+Shift+D)
'------------------------------------------------------------
Public Sub 파일경로복사()
    Dim s As String

    If ActiveWorkbook Is Nothing Then Exit Sub

    If ActiveWorkbook.Path = "" Then
        MsgBox "아직 저장되지 않은 파일입니다.", vbExclamation, "파일 경로 복사"
        Exit Sub
    End If

    s = ActiveWorkbook.FullName

    If 클립보드에넣기(s) Then
        MsgBox "파일 경로를 복사했습니다." & vbCrLf & vbCrLf & s, _
               vbInformation, "파일 경로 복사"
    Else
        InputBox "자동 복사에 실패했습니다. 아래 경로를 직접 복사(Ctrl+C)하세요.", _
                 "파일 경로 복사", s
    End If
End Sub

Private Function 클립보드에넣기(ByVal s As String) As Boolean
    Dim obj As Object
    On Error GoTo Fail
    ' MSForms.DataObject 를 참조 추가 없이 생성 (Late Binding)
    Set obj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    obj.SetText s
    obj.PutInClipboard
    클립보드에넣기 = True
    Exit Function
Fail:
    클립보드에넣기 = False
End Function


'============================================================
'  F. 업무도구 메뉴 (Ctrl+Shift+M)
'     UserForm 없이 CommandBar 팝업 - 참조 추가 불필요
'============================================================

Public Sub 매크로런처()
    Dim cb As CommandBar

    On Error Resume Next
    Application.CommandBars(메뉴이름).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=메뉴이름, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, _
                                         Temporary:=True)

    메뉴항목 cb, "새 창 열기 (같은 파일을 창 두 개로)", "새창열기", "Ctrl+Shift+O"
    메뉴항목 cb, "새 창으로 복사", "새창으로복사", "Ctrl+Shift+Q"
    메뉴항목 cb, "쉼표 형식", "쉼표형식", "Ctrl+Shift+X"
    메뉴항목 cb, "텍스트 나누기 - 일반으로", "일반으로변환", "Ctrl+Shift+Y"
    메뉴항목 cb, "텍스트 나누기 - 텍스트로", "텍스트로변환", "Ctrl+Shift+I"
    메뉴항목 cb, "표 테두리", "표테두리", "Ctrl+Shift+U"
    메뉴항목 cb, "행/열 자동맞춤", "행열자동맞춤", "Ctrl+Shift+W"
    메뉴항목 cb, "첫 행 고정 / 해제", "첫행고정", "Ctrl+Shift+S"
    메뉴항목 cb, "시트 일괄정리 (모든 시트 A1으로)", "시트일괄정리", "Ctrl+Shift+H"

    메뉴항목 cb, "피벗테이블 만들기 (새 시트)", "피벗만들기", "Ctrl+Shift+J", True
    메뉴항목 cb, "클래식 피벗 정리 (되돌리기 불가)", "클래식피벗", "Ctrl+Shift+T"
    메뉴항목 cb, "중복행 색칠", "중복행색칠", "Ctrl+Shift+D"
    메뉴항목 cb, "셀 배경 지우기 (범위 선택 필요)", "중복색지우기", "Ctrl+Shift+C"

    메뉴항목 cb, "필터 켜기 / 끄기", "자동필터", "Ctrl+Shift+A", True
    메뉴항목 cb, "선택한 셀 값으로 필터", "선택값으로필터", "Ctrl+Shift+Z"
    메뉴항목 cb, "고급필터 실행", "고급필터", "Ctrl+Shift+F"
    메뉴항목 cb, "고급필터 조건 열 세팅", "조건블록만들기", "Ctrl+Shift+R"
    메뉴항목 cb, "고급필터 범위 확인", "고급필터범위확인", ""
    메뉴항목 cb, "SUBTOTAL 합계", "소계넣기", "Ctrl+Shift+G"

    메뉴항목 cb, "시트 빼서 텍스트 저장 (txt→dat 변환용)", "시트빼서텍스트저장", "Ctrl+Shift+E", True
    메뉴항목 cb, "파일 경로 복사", "파일경로복사", "Ctrl+Shift+N"

    메뉴항목 cb, "암호 설정", "암호설정", "Ctrl+Shift+P", True
    메뉴항목 cb, 자동확인캡션(), "자동확인전환", ""
    메뉴항목 cb, "기본 암호 변경 (현재 : " & 기본암호읽기() & ")", "기본암호변경", ""

    메뉴항목 cb, "단축키 목록 붙이기 / 지우기", "치트시트", "Ctrl+Shift+K", True
    메뉴항목 cb, "설정 초기화 (배포 전 실행)", "설정초기화", ""
    메뉴항목 cb, "업무도구 다시 시작", "업무도구_다시시작", ""
    메뉴항목 cb, "진단 / 버전 정보", "업무도구_진단", ""

    cb.ShowPopup
    Exit Sub

Fail:
    MsgBox "메뉴를 표시하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "업무도구"
End Sub

Private Sub 메뉴항목(ByVal cb As CommandBar, _
                     ByVal 표시 As String, _
                     ByVal 매크로 As String, _
                     ByVal 키표시 As String, _
                     Optional ByVal 구분선 As Boolean = False)
    Dim btn As CommandBarButton

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    If Len(키표시) > 0 Then
        btn.Caption = 표시 & "   (" & 키표시 & ")"
    Else
        btn.Caption = 표시
    End If
    btn.OnAction = 접두어() & 매크로
    btn.BeginGroup = 구분선
    btn.Style = msoButtonCaption
End Sub

Private Sub 메뉴제목(ByVal cb As CommandBar, ByVal 표시 As String)
    Dim btn As CommandBarButton
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = 표시
    btn.Style = msoButtonCaption
    btn.Enabled = False
End Sub

Private Function 자동확인캡션() As String
    If 자동확인사용() Then
        자동확인캡션 = "닫을 때 암호 확인 : 켜짐  →  끄기"
    Else
        자동확인캡션 = "닫을 때 암호 확인 : 꺼짐  →  켜기"
    End If
End Function


'============================================================
'  G. 고급필터
'     데이터 블록 아래에 빈 줄을 두고 조건 블록을 두는 방식.
'     범위를 매번 지정할 필요 없이 두 블록을 자동으로 찾는다.
'============================================================

'------------------------------------------------------------
' 시트에서 첫 번째 데이터 블록 찾기
'------------------------------------------------------------
Private Function 데이터블록(ByVal ws As Worksheet) As Range
    Dim ur As Range
    Dim r As Long, c As Long
    Dim 끝행 As Long, 끝열 As Long

    On Error GoTo Fail
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function

    끝행 = ur.Row + ur.Rows.Count - 1
    끝열 = ur.Column + ur.Columns.Count - 1

    For r = ur.Row To 끝행
        If Application.CountA(ws.Rows(r)) > 0 Then
            For c = ur.Column To 끝열
                If Len(CStr(ws.Cells(r, c).Value)) > 0 Then
                    Set 데이터블록 = 소계행제외(ws.Cells(r, c).CurrentRegion)
                    Exit Function
                End If
            Next c
        End If
    Next r
    Exit Function
Fail:
End Function

'------------------------------------------------------------
' 블록 첫 행이 SUBTOTAL 합계 행이면 한 줄 아래를 머리글로 본다
'------------------------------------------------------------
Private Function 소계행제외(ByVal rg As Range) As Range
    Dim i As Long
    Dim 소계 As Boolean
    Dim f As String

    Set 소계행제외 = rg
    If rg Is Nothing Then Exit Function
    If rg.Rows.Count < 3 Then Exit Function

    On Error GoTo Done

    For i = 1 To rg.Columns.Count
        If rg.Cells(1, i).HasFormula Then
            f = UCase$(CStr(rg.Cells(1, i).Formula))
            If InStr(f, "SUBTOTAL") > 0 Then
                소계 = True
                Exit For
            End If
        End If
    Next i

    If 소계 Then
        Set 소계행제외 = rg.Offset(1, 0).Resize(rg.Rows.Count - 1, rg.Columns.Count)
    End If

Done:
End Function

'------------------------------------------------------------
' 데이터 블록 아래의 조건 블록 찾기 (없으면 Nothing)
'------------------------------------------------------------
Private Function 조건블록(ByVal ws As Worksheet, ByVal 데이터 As Range) As Range
    Dim ur As Range
    Dim r As Long, c As Long
    Dim 시작 As Long, 끝행 As Long, 끝열 As Long

    On Error GoTo Fail
    If 데이터 Is Nothing Then Exit Function

    Set ur = ws.UsedRange
    시작 = 데이터.Row + 데이터.Rows.Count
    끝행 = ur.Row + ur.Rows.Count - 1
    끝열 = ur.Column + ur.Columns.Count - 1
    If 시작 > 끝행 Then Exit Function

    For r = 시작 To 끝행
        If Application.CountA(ws.Rows(r)) > 0 Then
            For c = ur.Column To 끝열
                If Len(CStr(ws.Cells(r, c).Value)) > 0 Then
                    Set 조건블록 = ws.Cells(r, c).CurrentRegion
                    Exit Function
                End If
            Next c
        End If
    Next r
    Exit Function
Fail:
End Function

'------------------------------------------------------------
' 조건 머리글이 데이터 머리글에 다 있는지 확인
'------------------------------------------------------------
Private Function 머리글확인(ByVal 데이터 As Range, ByVal 조건 As Range, _
                            ByRef 잘못된이름 As String) As Boolean
    Dim i As Long, j As Long
    Dim h As String
    Dim 찾음 As Boolean

    잘못된이름 = ""
    On Error GoTo Fail

    For i = 1 To 조건.Columns.Count
        h = Trim$(CStr(조건.Cells(1, i).Value))
        If Len(h) > 0 Then
            찾음 = False
            For j = 1 To 데이터.Columns.Count
                If StrComp(h, Trim$(CStr(데이터.Cells(1, j).Value)), vbTextCompare) = 0 Then
                    찾음 = True
                    Exit For
                End If
            Next j
            If Not 찾음 Then
                잘못된이름 = h
                Exit Function
            End If
        End If
    Next i

    머리글확인 = True
    Exit Function
Fail:
    머리글확인 = True                       ' 판정 불가 시 통과시키고 Excel 오류에 맡김
End Function

'------------------------------------------------------------
' 고급필터 (Ctrl+Shift+F)
'   조건 블록이 없으면 만들어주고, 있으면 필터를 실행한다.
'------------------------------------------------------------
Public Sub 고급필터()
    Dim ws As Worksheet
    Dim 데이터 As Range, 조건 As Range
    Dim 잘못된이름 As String
    Dim n As Long

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "고급필터"
        Exit Sub
    End If
    Set ws = ActiveSheet

    '--- 이미 고급필터가 걸려 있으면 해제 (토글) ---
    On Error Resume Next
    If ws.FilterMode And Not ws.AutoFilterMode Then
        ws.ShowAllData
        Err.Clear
        On Error GoTo 0
        Application.StatusBar = "고급필터를 해제했습니다. 다시 누르면 조건대로 필터합니다."
        상태표시줄예약
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    Set 데이터 = 데이터블록(ws)
    If 데이터 Is Nothing Then
        MsgBox "데이터를 찾지 못했습니다.", vbExclamation, "고급필터"
        Exit Sub
    End If
    If 데이터.Rows.Count < 2 Then
        MsgBox "머리글과 데이터가 함께 있어야 합니다.", vbExclamation, "고급필터"
        Exit Sub
    End If

    Set 조건 = 조건블록(ws, 데이터)

    '--- 조건 블록이 없으면 만들어 준다 ---
    If 조건 Is Nothing Then
        조건블록만들기
        Exit Sub
    End If

    '--- 조건 값이 아직 없으면 입력 위치로 보낸다 ---
    If 조건.Rows.Count < 2 Then
        On Error Resume Next
        Application.Goto ws.Cells(조건.Row + 1, 조건.Column), False
        ws.Cells(조건.Row + 1, 조건.Column).Select
        Err.Clear
        On Error GoTo 0
        MsgBox "조건 블록에 아직 값이 없습니다." & vbCrLf & _
               "머리글 아래에 조건을 입력한 뒤 다시 실행해주세요.", _
               vbExclamation, "고급필터"
        Exit Sub
    End If

    '--- 머리글 유효성 ---
    If Not 머리글확인(데이터, 조건, 잘못된이름) Then
        MsgBox "조건 머리글 '" & 잘못된이름 & "' 이(가) 데이터 머리글에 없습니다." & vbCrLf & _
               "이름을 데이터와 똑같이 맞춰주세요.", vbExclamation, "고급필터"
        Exit Sub
    End If

    On Error GoTo Fail

    ' 이전 필터 상태를 먼저 해제
    On Error Resume Next
    ws.ShowAllData
    Err.Clear
    On Error GoTo Fail

    데이터.AdvancedFilter Action:=xlFilterInPlace, _
                          CriteriaRange:=조건, _
                          Unique:=False

    n = 표시된행수(데이터)

    Application.StatusBar = "고급필터 적용  |  목록 " & 데이터.Address(False, False) & _
                            "  조건 " & 조건.Address(False, False) & _
                            "  결과 " & Format$(n, "#,##0") & "건  |  해제 Ctrl+Shift+R"
    상태표시줄예약
    Exit Sub

Fail:
    MsgBox "고급필터를 적용하지 못했습니다." & vbCrLf & vbCrLf & _
           "목록 범위: " & 데이터.Address(False, False) & vbCrLf & _
           "조건 범위: " & 조건.Address(False, False) & vbCrLf & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "고급필터"
End Sub

'------------------------------------------------------------
' 필터 결과 건수 (머리글 제외)
'------------------------------------------------------------
Private Function 표시된행수(ByVal 데이터 As Range) As Long
    Dim rng As Range
    On Error GoTo Fail
    Set rng = 데이터.Columns(1).Offset(1, 0).Resize(데이터.Rows.Count - 1, 1)
    표시된행수 = rng.SpecialCells(xlCellTypeVisible).Cells.Count
    Exit Function
Fail:
    표시된행수 = 0
End Function

'------------------------------------------------------------
' 조건 블록 만들기 - 데이터 머리글을 아래에 복사
'------------------------------------------------------------
Public Sub 조건블록만들기()
    Dim errNo As Long, errDesc As String
    Dim ws As Worksheet
    Dim 데이터 As Range, 기존 As Range
    Dim 머리글 As Range, 대상 As Range
    Dim 되돌범위 As Range
    Dim 붙일행 As Long
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "조건 블록 만들기"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set 데이터 = 데이터블록(ws)
    If 데이터 Is Nothing Then
        MsgBox "데이터를 찾지 못했습니다.", vbExclamation, "조건 블록 만들기"
        Exit Sub
    End If

    Set 기존 = 조건블록(ws, 데이터)
    If Not 기존 Is Nothing Then
        If MsgBox("이미 조건 열이 있습니다. (" & 기존.Address(False, False) & ")" & vbCrLf & _
                  "지우고 새로 만들까요?" & vbCrLf & vbCrLf & _
                  "Ctrl+Z 로 되돌릴 수 있습니다.", _
                  vbYesNo + vbQuestion, "조건 열 세팅") = vbNo Then Exit Sub
    End If

    On Error GoTo Fail

    Set 머리글 = 데이터.Rows(1)
    붙일행 = 데이터.Row + 데이터.Rows.Count      ' 데이터 끝 + 1 = 빈 줄
    붙일행 = 붙일행 + 1                          ' 그 아래에 조건 머리글

    Set 대상 = ws.Cells(붙일행, 데이터.Column)

    '--- 되돌리기용 보관 : 지우기 전에, 기존 블록과 붙일 자리를 모두 ---
    r1 = 붙일행
    r2 = 붙일행
    c1 = 데이터.Column
    c2 = 데이터.Column + 데이터.Columns.Count - 1

    If Not 기존 Is Nothing Then
        If 기존.Row < r1 Then r1 = 기존.Row
        If 기존.Row + 기존.Rows.Count - 1 > r2 Then r2 = 기존.Row + 기존.Rows.Count - 1
        If 기존.Column < c1 Then c1 = 기존.Column
        If 기존.Column + 기존.Columns.Count - 1 > c2 Then c2 = 기존.Column + 기존.Columns.Count - 1
    End If

    Set 되돌범위 = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    조건열보관 되돌범위

    '--- 이제 지운다 ---
    If Not 기존 Is Nothing Then
        On Error Resume Next
        기존.Clear
        Err.Clear
        On Error GoTo Fail
    End If

    머리글.Copy
    대상.PasteSpecial xlPasteValues
    대상.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False

    On Error Resume Next
    Application.OnUndo "조건 열 세팅 취소", 접두어() & "조건열되돌리기"
    Err.Clear
    On Error GoTo Fail

    On Error Resume Next
    ws.Cells(붙일행 + 1, 데이터.Column).Select
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "조건 열을 세팅했습니다. 조건을 입력한 뒤 Ctrl+Shift+F 를 누르세요." & _
                            "  |  Ctrl+Z 로 되돌리기"
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    Err.Clear
    On Error GoTo 0
    MsgBox "조건 열을 세팅하지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "조건 열 세팅"
End Sub


'------------------------------------------------------------
' 고급필터 범위 확인 - 무엇을 데이터/조건으로 인식했는지 진단
'------------------------------------------------------------
Public Sub 고급필터범위확인()
    Dim ws As Worksheet
    Dim 데이터 As Range, 조건 As Range
    Dim s As String
    Dim 빈줄 As Long
    Dim 잘못된이름 As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "고급필터 범위 확인"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set 데이터 = 데이터블록(ws)

    If 데이터 Is Nothing Then
        MsgBox "데이터 블록을 찾지 못했습니다." & vbCrLf & _
               "시트가 비어 있는지 확인해주세요.", vbExclamation, "고급필터 범위 확인"
        Exit Sub
    End If

    s = "■ 데이터 블록" & vbCrLf & _
        "   범위 : " & 데이터.Address(False, False) & vbCrLf & _
        "   크기 : " & Format$(데이터.Rows.Count, "#,##0") & "행 × " & _
                       데이터.Columns.Count & "열" & vbCrLf & _
        "   머리글 : " & 머리글목록(데이터) & vbCrLf & vbCrLf

    Set 조건 = 조건블록(ws, 데이터)

    If 조건 Is Nothing Then
        빈줄 = 데이터.Row + 데이터.Rows.Count
        s = s & "■ 조건 블록  →  없음" & vbCrLf & vbCrLf & _
                "데이터가 " & (데이터.Row + 데이터.Rows.Count - 1) & "행에서 끝나고" & vbCrLf & _
                "그 아래에 값이 있는 행이 없습니다." & vbCrLf & vbCrLf & _
                "[조건 블록 새로 만들기]를 실행하면" & vbCrLf & _
                빈줄 & "행을 비우고 " & (빈줄 + 1) & "행에 머리글을 넣어드립니다."
        MsgBox s, vbInformation, "고급필터 범위 확인"
        Exit Sub
    End If

    s = s & "■ 조건 블록" & vbCrLf & _
            "   범위 : " & 조건.Address(False, False) & vbCrLf & _
            "   크기 : " & 조건.Rows.Count & "행 × " & 조건.Columns.Count & "열" & vbCrLf & _
            "   머리글 : " & 머리글목록(조건) & vbCrLf & vbCrLf

    If 조건.Rows.Count < 2 Then
        s = s & "△ 조건 값이 아직 없습니다." & vbCrLf & _
                "   머리글 아래 " & (조건.Row + 1) & "행에 조건을 입력하세요."
    ElseIf Not 머리글확인(데이터, 조건, 잘못된이름) Then
        s = s & "△ 조건 머리글 '" & 잘못된이름 & "' 이(가)" & vbCrLf & _
                "   데이터 머리글에 없습니다. 철자를 맞춰주세요."
    Else
        s = s & "○ 바로 필터할 수 있습니다.  (Ctrl+Shift+F)"
    End If

    MsgBox s, vbInformation, "고급필터 범위 확인"
End Sub

Private Function 머리글목록(ByVal rng As Range) As String
    Dim i As Long
    Dim s As String
    Dim v As String

    On Error GoTo Fail
    For i = 1 To rng.Columns.Count
        v = Trim$(CStr(rng.Cells(1, i).Value))
        If Len(v) = 0 Then v = "(빈칸)"
        If Len(s) > 0 Then s = s & " | "
        s = s & v
        If i >= 8 Then
            s = s & " | ..."
            Exit For
        End If
    Next i
    머리글목록 = s
    Exit Function
Fail:
    머리글목록 = "(확인 실패)"
End Function

'------------------------------------------------------------
' 자동 필터 켜기 / 끄기 (Ctrl+Shift+A)
'   표(ListObject) 안이면 표의 필터 단추를 토글한다.
'------------------------------------------------------------
Public Sub 자동필터()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lo As Object
    Dim 자동단추 As Boolean
    Dim 걸림 As Boolean

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "필터"
        Exit Sub
    End If
    Set ws = ActiveSheet

    On Error GoTo Fail

    '--- 표(ListObject) 안이면 표의 필터를 토글 ---
    On Error Resume Next
    Set lo = ActiveCell.ListObject
    Err.Clear
    On Error GoTo Fail

    If Not lo Is Nothing Then
        lo.ShowAutoFilter = Not lo.ShowAutoFilter
        Application.StatusBar = "표 '" & lo.Name & "' 의 필터를 " & _
                                IIf(lo.ShowAutoFilter, "켰습니다.", "껐습니다.")
        상태표시줄예약
        Exit Sub
    End If

    On Error Resume Next
    자동단추 = ws.AutoFilterMode
    걸림 = ws.FilterMode
    Err.Clear
    On Error GoTo Fail

    '--- 고급필터로 행이 숨겨진 상태면 그것부터 해제 ---
    If 걸림 And Not 자동단추 Then
        On Error Resume Next
        ws.ShowAllData
        Err.Clear
        On Error GoTo Fail
        Application.StatusBar = "고급필터를 해제했습니다. 모든 행이 표시됩니다."
        상태표시줄예약
        Exit Sub
    End If

    '--- 자동 필터가 있으면 끄기 (리본의 [필터] 단추와 동일) ---
    If 자동단추 Then
        On Error Resume Next
        If 걸림 Then ws.ShowAllData
        Err.Clear
        On Error GoTo Fail
        ws.AutoFilterMode = False
        Application.StatusBar = "필터를 껐습니다."
        상태표시줄예약
        Exit Sub
    End If

    '--- 없으면 걸기 ---
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.CountLarge >= 2 Then
            Set rng = 소계행제외(Selection)
        Else
            Set rng = 소계행제외(ActiveCell.CurrentRegion)
        End If
    End If

    If rng Is Nothing Then
        MsgBox "필터를 걸 범위를 찾지 못했습니다.", vbExclamation, "필터"
        Exit Sub
    End If
    If rng.Cells.CountLarge < 2 Then
        MsgBox "데이터가 있는 셀에 커서를 두고 실행해주세요.", vbExclamation, "필터"
        Exit Sub
    End If

    rng.AutoFilter

    Application.StatusBar = "필터를 켰습니다. (" & rng.Address(False, False) & ")"
    상태표시줄예약
    Exit Sub

Fail:
    MsgBox "필터를 바꾸지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "필터"
End Sub

'------------------------------------------------------------
' 첫 행 고정 / 해제 (Ctrl+Shift+S)
'   머리글 행을 찾아 그 아래에서 틀을 고정한다. 다시 누르면 해제.
'------------------------------------------------------------
Public Sub 첫행고정()
    Dim ws As Worksheet
    Dim 데이터 As Range
    Dim 머리글행 As Long
    Dim 되돌릴셀 As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "첫 행 고정"
        Exit Sub
    End If
    Set ws = ActiveSheet

    On Error GoTo Fail

    If ActiveWindow.FreezePanes Then
        ActiveWindow.FreezePanes = False
        Application.StatusBar = "틀 고정을 해제했습니다."
        상태표시줄예약
        Exit Sub
    End If

    '--- 머리글 행 찾기 ---
    머리글행 = 1
    Set 데이터 = 데이터블록(ws)
    If Not 데이터 Is Nothing Then 머리글행 = 데이터.Row

    If TypeName(Selection) = "Range" Then Set 되돌릴셀 = Selection

    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Err.Clear
    On Error GoTo Fail

    ws.Cells(머리글행 + 1, 1).Select
    ActiveWindow.FreezePanes = True

    On Error Resume Next
    If Not 되돌릴셀 Is Nothing Then 되돌릴셀.Select
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = 머리글행 & "행까지 고정했습니다. 다시 누르면 해제됩니다."
    상태표시줄예약
    Exit Sub

Fail:
    MsgBox "틀 고정을 바꾸지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "첫 행 고정"
End Sub

'------------------------------------------------------------
' SUBTOTAL 합계 (Ctrl+Shift+G)
'   커서 셀 바로 아래는 머리글로 보고 건너뛰고,
'   그 다음 행부터 데이터 끝까지를 SUBTOTAL(9,) 로 합산한다.
'   가로로 여러 칸을 선택하면 각 열마다 따로 넣는다.
'------------------------------------------------------------
Public Sub 소계넣기()
    Dim ws As Worksheet
    Dim sel As Range
    Dim c As Range
    Dim 시작행 As Long, 끝행 As Long
    Dim 넣은수 As Long
    Dim oldScreen As Boolean
    Dim 마지막 As Long
    Dim i As Long
    Dim 이전수식() As Variant
    Dim 이전굵게() As Variant

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "SUBTOTAL 합계"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "합계를 넣을 셀을 선택한 뒤 실행해주세요.", vbExclamation, "SUBTOTAL 합계"
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "떨어져 있는 여러 범위는 한 번에 처리할 수 없습니다.", vbExclamation, "SUBTOTAL 합계"
        Exit Sub
    End If
    If sel.Rows.Count > 1 Then
        MsgBox "한 행만 선택해주세요." & vbCrLf & _
               "합계를 넣을 칸들을 가로로 선택하면 각 열마다 넣어드립니다.", _
               vbExclamation, "SUBTOTAL 합계"
        Exit Sub
    End If

    ' 머리글 한 줄 건너뛴 다음 행부터
    시작행 = sel.Row + 2
    If 시작행 > ws.Rows.Count Then
        MsgBox "데이터를 찾을 수 없습니다.", vbExclamation, "SUBTOTAL 합계"
        Exit Sub
    End If

    ' 데이터 끝 행 : 머리글 아래 표 블록의 마지막 행
    '   빈 행으로 떨어져 있는 조건 블록은 자동으로 제외된다
    끝행 = 영역의마지막행(ws, sel.Cells(1, 1).Offset(1, 0))

    If 끝행 < 시작행 Then
        끝행 = 영역의마지막행(ws, sel.Cells(1, 1).Offset(2, 0))
    End If

    ' 표 블록을 못 찾으면 열 단위로 대체 (조건 블록이 없는 단순 표)
    If 끝행 < 시작행 Then
        For Each c In sel.Cells
            마지막 = 열의마지막행(ws, c.Column, 시작행)
            If 마지막 > 끝행 Then 끝행 = 마지막
        Next c
    End If

    If 끝행 < 시작행 Then
        MsgBox "합계를 낼 데이터가 없습니다." & vbCrLf & vbCrLf & _
               "합계를 넣을 칸에 커서를 두고 실행하세요." & vbCrLf & _
               "바로 아래 칸은 머리글로 보고 건너뜁니다.", vbExclamation, "SUBTOTAL 합계"
        Exit Sub
    End If

    '--- 마지막으로 G 를 누른 자리와 같으면 원래대로 되돌린다 ---
    If mUndoOn And Len(mUndoAddr) > 0 Then
        If StrComp(mUndoWb, ws.Parent.Name, vbTextCompare) = 0 And _
           StrComp(mUndoWs, ws.Name, vbTextCompare) = 0 And _
           StrComp(mUndoAddr, sel.Address, vbTextCompare) = 0 Then
            소계되돌리기
            Exit Sub
        End If
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    ' 되돌리기용으로 기존 값 보관
    ReDim 이전수식(1 To sel.Cells.Count)
    ReDim 이전굵게(1 To sel.Cells.Count)
    i = 0
    For Each c In sel.Cells
        i = i + 1
        이전수식(i) = c.Formula
        이전굵게(i) = c.Font.Bold
    Next c
    mUndoWb = ws.Parent.Name
    mUndoWs = ws.Name
    mUndoAddr = sel.Address
    mUndoFormula = 이전수식
    mUndoBold = 이전굵게
    mUndoOn = True

    For Each c In sel.Cells
        c.Formula = "=SUBTOTAL(9," & ws.Cells(시작행, c.Column).Address(False, False) & ":" & _
                                     ws.Cells(끝행, c.Column).Address(False, False) & ")"
        c.Font.Bold = True
        넣은수 = 넣은수 + 1
    Next c

    Application.ScreenUpdating = oldScreen

    ' Ctrl+Z 로 되돌릴 수 있게 등록
    On Error Resume Next
    Application.OnUndo "SUBTOTAL 합계 취소", 접두어() & "소계되돌리기"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "SUBTOTAL 합계 " & 넣은수 & "칸을 넣었습니다. (" & _
                            시작행 & "행 ~ " & 끝행 & "행)  |  " & _
                            "Ctrl+Z 또는 Ctrl+Shift+G 로 되돌리기"
    상태표시줄예약
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "합계를 넣지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "SUBTOTAL 합계"
End Sub

'--- 해당 열의 마지막 데이터 행 (아래에서 위로 거슬러 올라감) ---
Private Function 열의마지막행(ByVal ws As Worksheet, _
                               ByVal 열번호 As Long, _
                               ByVal 시작행 As Long) As Long
    Dim r As Long
    On Error GoTo Fail
    r = ws.Cells(ws.Rows.Count, 열번호).End(xlUp).Row
    If r < 시작행 Then r = 0
    열의마지막행 = r
    Exit Function
Fail:
    열의마지막행 = 0
End Function

'--- 표 전체(CurrentRegion) 기준 마지막 행 ---
Private Function 영역의마지막행(ByVal ws As Worksheet, ByVal 기준셀 As Range) As Long
    Dim rg As Range
    On Error GoTo Fail
    Set rg = 기준셀.CurrentRegion
    If rg Is Nothing Then Exit Function
    영역의마지막행 = rg.Row + rg.Rows.Count - 1
    Exit Function
Fail:
    영역의마지막행 = 0
End Function

'------------------------------------------------------------
' SUBTOTAL 합계 되돌리기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 소계되돌리기()
    Dim ws As Worksheet
    Dim rg As Range
    Dim i As Long
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If Len(mUndoAddr) = 0 Then
        MsgBox "되돌릴 SUBTOTAL 합계가 없습니다.", vbInformation, "SUBTOTAL 합계"
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
    Next i

    mUndoAddr = ""
    mUndoOn = False
    Application.ScreenUpdating = oldScreen

    Application.StatusBar = "SUBTOTAL 합계를 되돌렸습니다."
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "되돌리지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "SUBTOTAL 합계"
End Sub

'------------------------------------------------------------
' 중복행 색칠 되돌리기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 중복색되돌리기()
    Dim ws As Worksheet
    Dim rng As Range
    Dim r As Long, c As Long
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If Len(mDupAddr) = 0 Or mDupCount = 0 Then
        MsgBox "되돌릴 중복 색칠이 없습니다.", vbInformation, "중복행 색칠"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set ws = Application.Workbooks(mDupWb).Worksheets(mDupWs)
    Set rng = ws.Range(mDupAddr)

    For r = 1 To mDupCount
        For c = 1 To UBound(mDupColor, 2)
            With rng.Cells(mDupRows(r), c).Interior
                If mDupIndex(r, c) = xlColorIndexNone Then
                    .ColorIndex = xlColorIndexNone
                Else
                    .Color = mDupColor(r, c)
                End If
            End With
        Next c
    Next r

    mDupAddr = ""
    mDupCount = 0
    Application.ScreenUpdating = oldScreen

    Application.StatusBar = "중복행 색칠을 되돌렸습니다."
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "되돌리지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "중복행 색칠"
End Sub

'------------------------------------------------------------
' 선택한 셀 값으로 필터 (Ctrl+Shift+X)
'   필터가 없으면 현재 영역에 걸고, 커서 셀 값으로 조건을 잡는다.
'------------------------------------------------------------
Public Sub 선택값으로필터()
    Dim ws As Worksheet
    Dim lo As Object
    Dim fr As Range
    Dim 기준 As Range
    Dim c As Range
    Dim v As Variant
    Dim crit As String
    Dim fld As Long
    Dim 남은행 As Long
    Dim 걸림 As Boolean
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "선택 값으로 필터"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "필터를 걸 셀을 선택한 뒤 실행해주세요.", vbExclamation, "선택 값으로 필터"
        Exit Sub
    End If
    Set c = ActiveCell

    On Error GoTo Fail

    '--- 표(ListObject) 안이면 표 기준 ---
    On Error Resume Next
    Set lo = c.ListObject
    Err.Clear
    On Error GoTo Fail

    If Not lo Is Nothing Then
        If Not lo.ShowAutoFilter Then lo.ShowAutoFilter = True
        Set fr = lo.Range
    Else
        Set 기준 = 소계행제외(c.CurrentRegion)
        If 기준 Is Nothing Then
            MsgBox "데이터를 찾지 못했습니다.", vbExclamation, "선택 값으로 필터"
            Exit Sub
        End If
        If 기준.Cells.CountLarge < 2 Or 기준.Rows.Count < 2 Then
            MsgBox "데이터가 있는 셀에 커서를 두고 실행해주세요.", _
                   vbExclamation, "선택 값으로 필터"
            Exit Sub
        End If

        If ws.AutoFilterMode Then
            ' 기존 필터의 머리글 위치가 어긋나 있으면 새로 건다
            If ws.AutoFilter.Range.Row <> 기준.Row Or _
               ws.AutoFilter.Range.Column <> 기준.Column Then
                On Error Resume Next
                ws.ShowAllData
                Err.Clear
                On Error GoTo Fail
                ws.AutoFilterMode = False
                기준.AutoFilter
            End If
        Else
            기준.AutoFilter
        End If

        Set fr = ws.AutoFilter.Range
    End If

    fld = c.Column - fr.Column + 1
    If fld < 1 Or fld > fr.Columns.Count Then
        MsgBox "필터 범위 안의 셀을 선택해주세요." & vbCrLf & _
               "필터 범위: " & fr.Address(False, False), vbExclamation, "선택 값으로 필터"
        Exit Sub
    End If
    If c.Row = fr.Row Then
        MsgBox "머리글이 아니라 데이터 셀을 선택해주세요.", vbExclamation, "선택 값으로 필터"
        Exit Sub
    End If

    '--- 이미 이 열에 필터가 걸려 있으면 해제 (토글) ---
    걸림 = False
    On Error Resume Next
    걸림 = ws.AutoFilter.Filters(fld).On
    Err.Clear
    On Error GoTo Fail

    If 걸림 Then
        On Error Resume Next
        fr.AutoFilter Field:=fld
        Err.Clear
        On Error GoTo Fail
        Application.StatusBar = fr.Cells(1, fld).Value & " 열의 필터를 해제했습니다.  " & _
                                "|  전체 해제 Ctrl+Shift+A"
        상태표시줄예약
        Exit Sub
    End If

    v = c.Value

    '--- 1차 시도 ---
    If IsEmpty(v) Or Len(CStr(v)) = 0 Then
        crit = "(빈 셀)"
        fr.AutoFilter Field:=fld, Criteria1:="="
    ElseIf IsDate(v) And Not IsNumeric(v) Then
        crit = Format$(v, "yyyy-mm-dd")
        fr.AutoFilter Field:=fld, _
                      Criteria1:=">=" & CDbl(CDate(v)), _
                      Operator:=xlAnd, _
                      Criteria2:="<" & (CDbl(CDate(v)) + 1)
    ElseIf IsNumeric(v) Then
        crit = CStr(v)
        fr.AutoFilter Field:=fld, Criteria1:="=" & CStr(v)
    Else
        crit = CStr(v)
        fr.AutoFilter Field:=fld, Criteria1:="=" & 와일드카드보호(crit)
    End If

    남은행 = 표시된행수(fr)

    '--- 2차 시도 : 숫자는 표기 방식을 바꿔 다시 ---
    If 남은행 = 0 And IsNumeric(v) And Not IsEmpty(v) Then
        On Error Resume Next
        fr.AutoFilter Field:=fld, Criteria1:=v
        Err.Clear
        On Error GoTo Fail
        남은행 = 표시된행수(fr)

        If 남은행 = 0 Then
            On Error Resume Next
            fr.AutoFilter Field:=fld, Criteria1:=c.Text
            Err.Clear
            On Error GoTo Fail
            남은행 = 표시된행수(fr)
        End If
    End If

    '--- 그래도 0건이면 이 열의 조건만 풀어 원래대로 ---
    If 남은행 = 0 Then
        On Error Resume Next
        fr.AutoFilter Field:=fld
        Err.Clear
        On Error GoTo Fail

        MsgBox "조건에 맞는 행이 없어 이 열의 필터를 해제했습니다." & vbCrLf & vbCrLf & _
               "머리글: " & fr.Cells(1, fld).Value & vbCrLf & _
               "조건: " & crit & vbCrLf & _
               "필터 범위: " & fr.Address(False, False), _
               vbExclamation, "선택 값으로 필터"
        Exit Sub
    End If

    Application.StatusBar = fr.Cells(1, fld).Value & " = " & crit & "  ->  " & _
                            Format$(남은행, "#,##0") & "건  |  해제 Ctrl+Shift+A"
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "필터를 걸지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "선택 값으로 필터"
End Sub

'--- 필터 조건에서 * ? ~ 를 글자 그대로 취급하게 ---
Private Function 와일드카드보호(ByVal s As String) As String
    s = Replace(s, "~", "~~")
    s = Replace(s, "*", "~*")
    s = Replace(s, "?", "~?")
    와일드카드보호 = s
End Function

'------------------------------------------------------------
' 단축키 목록 붙이기 / 지우기 (Ctrl+Shift+K)
'   화면 오른쪽에 떠 있는 메모지처럼 붙인다. 다시 누르면 지워진다.
'------------------------------------------------------------
Public Sub 치트시트()
    Dim ws As Worksheet
    Dim shp As Object
    Dim 있음 As Boolean
    Dim 왼쪽 As Single, 위쪽 As Single
    Dim vr As Range
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "단축키 목록"
        Exit Sub
    End If
    Set ws = ActiveSheet

    On Error GoTo Fail

    ' 이미 있으면 지우기
    On Error Resume Next
    Set shp = ws.Shapes(치트시트이름)
    Err.Clear
    On Error GoTo Fail

    If Not shp Is Nothing Then
        shp.Delete
        Application.StatusBar = "단축키 목록을 지웠습니다."
        상태표시줄예약
        Exit Sub
    End If

    ' 화면 오른쪽 위에 붙인다
    왼쪽 = 40
    위쪽 = 40
    On Error Resume Next
    Set vr = ActiveWindow.VisibleRange
    If Not vr Is Nothing Then
        왼쪽 = vr.Left + vr.Width - 265
        위쪽 = vr.Top + 12
        If 왼쪽 < vr.Left Then 왼쪽 = vr.Left + 8
    End If
    Err.Clear
    On Error GoTo Fail

    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 왼쪽, 위쪽, 300, 200)

    With shp
        .Name = 치트시트이름
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
                .Text = 치트시트내용()
                .Font.Name = "맑은 고딕"
                .Font.Size = 9
                .Font.Color = RGB(35, 35, 40)
            End With
        End With
    End With

    ' 제목 줄 강조
    On Error Resume Next
    With shp.TextFrame.Characters(1, 14).Font
        .Bold = True
        .Size = 11
        .Color = RGB(0, 70, 150)
    End With
    Err.Clear
    On Error GoTo Fail

    단축키강조 shp

    ' 글자에 맞춰 상자 크기 자동 조정 (잘림 방지)
    On Error Resume Next
    With shp.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
    End With
    Err.Clear
    On Error GoTo Fail

    ' 화면 밖으로 나가지 않게 위치 보정
    On Error Resume Next
    If Not vr Is Nothing Then
        shp.Left = vr.Left + vr.Width - shp.Width - 12
        If shp.Left < vr.Left + 6 Then shp.Left = vr.Left + 6
        shp.Top = vr.Top + 6
    End If
    Err.Clear
    On Error GoTo Fail

    ' Ctrl+Z 로도 지울 수 있게
    mCheatWb = ws.Parent.Name
    mCheatWs = ws.Name
    On Error Resume Next
    Application.OnUndo "단축키 목록 지우기", 접두어() & "치트시트지우기"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "단축키 목록을 붙였습니다. 다시 Ctrl+Shift+K 를 누르면 지워집니다."
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "단축키 목록을 붙이지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "단축키 목록"
End Sub

Private Function 치트시트내용() As String
    Dim s As String
    s = "Excel 업무도구  v" & 도구버전 & vbLf & vbLf
    s = s & "Ctrl+Shift+M    업무도구 메뉴  -  기억 안 나면 이것만 누르세요" & vbLf & vbLf
    s = s & "[ 편집 ]" & vbLf
    s = s & "Ctrl+Shift+O    새 창 열기  (같은 파일을 창 두 개로)" & vbLf
    s = s & "Ctrl+Shift+Q    새 창으로 복사  (현재 시트를 새 통합문서로)" & vbLf
    s = s & "Ctrl+Shift+X    쉼표 형식  (음수는 빨강 △, 0은 -)" & vbLf
    s = s & "Ctrl+Shift+Y    텍스트 나누기 - 일반으로  (텍스트 숫자를 숫자로)" & vbLf
    s = s & "Ctrl+Shift+I    텍스트 나누기 - 텍스트로" & vbLf
    s = s & "Ctrl+Shift+U    표 테두리  (머리글 굵게 + 세로 구분선)" & vbLf
    s = s & "Ctrl+Shift+W    행/열 자동맞춤" & vbLf
    s = s & "Ctrl+Shift+S    첫 행 고정 / 해제  (머리글 자동 인식)" & vbLf
    s = s & "Ctrl+Shift+H    시트 일괄정리  (모든 시트 A1으로, 확대 100%)" & vbLf & vbLf
    s = s & "[ 데이터 ]" & vbLf
    s = s & "Ctrl+Shift+J    피벗테이블 만들기  (새 시트, 범위 자동)" & vbLf
    s = s & "Ctrl+Shift+T    클래식 피벗 정리  (부분합 제거, 되돌리기 불가)" & vbLf
    s = s & "Ctrl+Shift+D    중복행 색칠  (같은 값끼리 같은 색)" & vbLf
    s = s & "Ctrl+Shift+C    셀 배경 지우기  (선택 범위의 채우기 색 제거)" & vbLf
    s = s & "Ctrl+Shift+G    SUBTOTAL 합계  (머리글 위 칸에서, 다시 누르면 취소)" & vbLf & vbLf
    s = s & "[ 필터 ]" & vbLf
    s = s & "Ctrl+Shift+A    필터 켜기 / 끄기" & vbLf
    s = s & "Ctrl+Shift+Z    선택한 셀 값으로 필터  (다시 누르면 그 열만 해제)" & vbLf
    s = s & "Ctrl+Shift+F    고급필터 실행  (다시 누르면 해제)" & vbLf
    s = s & "Ctrl+Shift+R    고급필터 조건 열 세팅  (머리글 자동 복사)" & vbLf
    s = s & "        고급필터는 데이터 아래에 빈 줄을 두고" & vbLf
    s = s & "        조건 블록을 만들어 사용합니다." & vbLf & vbLf
    s = s & "[ 내보내기 ]" & vbLf
    s = s & "Ctrl+Shift+E    시트 빼서 텍스트 저장  (txt → dat 변환용)" & vbLf
    s = s & "Ctrl+Shift+N    파일 경로 복사  (전체 경로를 클립보드로)" & vbLf & vbLf
    s = s & "[ 기타 ]" & vbLf
    s = s & "Ctrl+Shift+P    암호 설정  (열기 암호, 기본값 1234)" & vbLf
    s = s & "Ctrl+Shift+K    이 목록 붙이기 / 지우기" & vbLf & vbLf
    s = s & "저장된 파일을 닫을 때 열기 암호가 없으면 자동으로 묻습니다."
    치트시트내용 = s
End Function

'------------------------------------------------------------
' 단축키 목록 지우기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 치트시트지우기()
    Dim ws As Worksheet
    Dim shp As Object

    If Len(mCheatWs) = 0 Then Exit Sub

    On Error Resume Next
    Set ws = Application.Workbooks(mCheatWb).Worksheets(mCheatWs)
    If Not ws Is Nothing Then
        Set shp = ws.Shapes(치트시트이름)
        If Not shp Is Nothing Then shp.Delete
    End If
    Err.Clear
    On Error GoTo 0

    mCheatWs = ""
    Application.StatusBar = "단축키 목록을 지웠습니다."
    상태표시줄예약
End Sub

'------------------------------------------------------------
' 텍스트 나누기 - 일반으로 (Ctrl+Shift+Y)
' 텍스트 나누기 - 텍스트로 (Ctrl+Shift+I)
'   구분자를 모두 끄고 서식만 다시 읽는다. 열이 쪼개지지 않는다.
'   여러 열을 한 번에 처리할 수 있다.
'------------------------------------------------------------
Public Sub 일반으로변환()
    텍스트나누기 xlGeneralFormat, "일반"
End Sub

Public Sub 텍스트로변환()
    텍스트나누기 xlTextFormat, "텍스트"
End Sub

Private Sub 텍스트나누기(ByVal 서식 As Long, ByVal 서식명 As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim blk As Range
    Dim col As Range
    Dim 안내 As String
    Dim 되돌리기가능 As Boolean
    Dim oldScreen As Boolean, oldEvents As Boolean, oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "텍스트 나누기"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "변환할 셀을 선택한 뒤 실행해주세요.", vbExclamation, "텍스트 나누기"
        Exit Sub
    End If

    '--- 범위 결정 ---
    If Selection.Cells.CountLarge >= 2 Then
        Set rng = Selection
    Else
        Set blk = 소계행제외(ActiveCell.CurrentRegion)
        If blk Is Nothing Then
            MsgBox "데이터를 찾지 못했습니다.", vbExclamation, "텍스트 나누기"
            Exit Sub
        End If
        Set rng = ws.Range(ws.Cells(blk.Row, ActiveCell.Column), _
                           ws.Cells(blk.Row + blk.Rows.Count - 1, ActiveCell.Column))
    End If

    If rng.Areas.Count > 1 Then
        MsgBox "떨어져 있는 여러 범위는 한 번에 처리할 수 없습니다." & vbCrLf & _
               "연속된 범위만 선택한 뒤 다시 실행해주세요.", vbExclamation, "텍스트 나누기"
        Exit Sub
    End If

    ' 전체 열/행 선택 방어
    On Error Resume Next
    Set rng = Application.Intersect(rng, ws.UsedRange)
    Err.Clear
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "변환할 데이터가 없습니다.", vbExclamation, "텍스트 나누기"
        Exit Sub
    End If

    되돌리기가능 = (rng.Cells.CountLarge <= 100000)

    '--- 실행 전 확인 ---
    안내 = "선택 범위를 [" & 서식명 & "] 형식으로 다시 읽습니다." & vbCrLf & vbCrLf & _
           "범위 : " & rng.Address(False, False) & _
           "   (" & Format$(rng.Columns.Count, "#,##0") & "개 열 / " & _
           Format$(rng.Rows.Count, "#,##0") & "행)" & vbCrLf & vbCrLf & _
           "· 구분자를 모두 꺼서 열이 쪼개지지 않습니다." & vbCrLf

    If 서식 = xlGeneralFormat Then
        안내 = 안내 & "· 텍스트로 저장된 숫자가 숫자로 바뀝니다." & vbCrLf & _
                     "· 앞자리 0 은 사라집니다.  (00123 → 123)" & vbCrLf
    Else
        안내 = 안내 & "· 모든 값이 텍스트가 됩니다. 계산에 쓸 수 없게 됩니다." & vbCrLf
    End If

    안내 = 안내 & IIf(되돌리기가능, "· Ctrl+Z 로 되돌릴 수 있습니다.", _
                                    "· 범위가 커서 Ctrl+Z 되돌리기는 지원되지 않습니다.") & _
           vbCrLf & vbCrLf & "계속할까요?"

    If MsgBox(안내, vbOKCancel + vbQuestion, "텍스트 나누기 - " & 서식명) = vbCancel Then Exit Sub

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    '--- 되돌리기용 보관 ---
    If 되돌리기가능 Then
        mTxtWb = ws.Parent.Name
        mTxtWs = ws.Name
        mTxtAddr = rng.Address
        mTxtValue = rng.Formula
        mTxtFormat = rng.NumberFormat
        mTxtOn = True
    Else
        mTxtOn = False
    End If

    '--- 열마다 서식만 다시 적용 ---
    For Each col In rng.Columns
        col.TextToColumns Destination:=col.Cells(1, 1), _
                          DataType:=xlDelimited, _
                          TextQualifier:=xlDoubleQuote, _
                          ConsecutiveDelimiter:=False, _
                          Tab:=False, Semicolon:=False, Comma:=False, _
                          Space:=False, Other:=False, _
                          FieldInfo:=Array(1, 서식), _
                          TrailingMinusNumbers:=True
    Next col

    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    If 되돌리기가능 Then
        On Error Resume Next
        Application.OnUndo "텍스트 나누기 취소", 접두어() & "텍스트나누기되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "[" & 서식명 & "] 형식으로 다시 읽었습니다.  " & _
                            rng.Address(False, False) & _
                            IIf(되돌리기가능, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "변환하지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "텍스트 나누기"
End Sub

'------------------------------------------------------------
' 텍스트 나누기 되돌리기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 텍스트나누기되돌리기()
    Dim ws As Worksheet
    Dim rng As Range
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim errNo As Long, errDesc As String

    If Not mTxtOn Or Len(mTxtAddr) = 0 Then
        MsgBox "되돌릴 텍스트 나누기가 없습니다.", vbInformation, "텍스트 나누기"
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

    Application.StatusBar = "텍스트 나누기를 되돌렸습니다."
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    MsgBox "되돌리지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "텍스트 나누기"
End Sub

'--- 치트시트에서 Ctrl+Shift+? 부분만 굵게 ---
Private Sub 단축키강조(ByVal shp As Object)
    Dim s As String
    Dim p As Long
    Dim 길이 As Long

    On Error Resume Next
    s = shp.TextFrame.Characters.Text
    길이 = Len("Ctrl+Shift+X")

    p = InStr(1, s, "Ctrl+Shift+", vbTextCompare)
    Do While p > 0
        With shp.TextFrame.Characters(p, 길이).Font
            .Bold = True
            .Color = RGB(0, 70, 150)
        End With
        p = InStr(p + 1, s, "Ctrl+Shift+", vbTextCompare)
    Loop

    ' 맨 위 메뉴 안내 줄 전체를 강조
    p = InStr(1, s, "Ctrl+Shift+M", vbTextCompare)
    If p > 0 Then
        shp.TextFrame.Characters(p, 50).Font.Bold = True
    End If

    Err.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------
' 되돌리기 - 쉼표 형식
'------------------------------------------------------------
Public Sub 쉼표형식되돌리기()
    On Error Resume Next
    If Not mFmtOn Or Len(mFmtAddr) = 0 Then
        MsgBox "되돌릴 쉼표 형식이 없습니다.", vbInformation, "쉼표 형식"
        Exit Sub
    End If
    Application.Workbooks(mFmtWb).Worksheets(mFmtWs).Range(mFmtAddr).NumberFormat = mFmtValue
    mFmtOn = False
    mFmtAddr = ""
    Err.Clear
    On Error GoTo 0
    Application.StatusBar = "쉼표 형식을 되돌렸습니다."
    상태표시줄예약
End Sub

'------------------------------------------------------------
' 되돌리기 - 행/열 자동맞춤
'------------------------------------------------------------
Public Sub 자동맞춤되돌리기()
    Dim rng As Range
    Dim i As Long
    Dim oldScreen As Boolean

    If Not mFitOn Or Len(mFitAddr) = 0 Then
        MsgBox "되돌릴 자동맞춤이 없습니다.", vbInformation, "행열 자동맞춤"
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

    Application.StatusBar = "행/열 너비를 되돌렸습니다."
    상태표시줄예약
End Sub

'------------------------------------------------------------
' 되돌리기 - 고급필터 조건 열 세팅
'------------------------------------------------------------
Public Sub 조건열되돌리기()
    Dim rng As Range
    Dim r As Long, c As Long
    Dim oldScreen As Boolean

    If Not mCondOn Or Len(mCondAddr) = 0 Then
        MsgBox "되돌릴 조건 열 세팅이 없습니다.", vbInformation, "조건 열 세팅"
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

    Application.StatusBar = "조건 열 세팅을 되돌렸습니다."
    상태표시줄예약
End Sub

'--- 조건 열 세팅 되돌리기용 : 값 + 서식 + 색을 통째로 보관 ---
Private Sub 조건열보관(ByVal rg As Range)
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
' 새 창 열기 (Ctrl+Shift+O)
'   같은 파일을 창 두 개로 띄운다. (보기 - 새 창 과 동일)
'   닫을 때는 그냥 창을 닫으면 된다.
'------------------------------------------------------------
Public Sub 새창열기()
    Dim wb As Workbook
    Dim 창수 As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then
        MsgBox "먼저 파일을 열어주세요.", vbExclamation, "새 창 열기"
        Exit Sub
    End If

    Set wb = ActiveWorkbook

    On Error GoTo Fail
    ActiveWindow.NewWindow

    On Error Resume Next
    창수 = wb.Windows.Count
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "새 창을 열었습니다.  (" & wb.Name & " 의 창 " & 창수 & "개)" & _
                            "   닫을 때는 창을 그냥 닫으면 됩니다."
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "새 창을 열지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "새 창 열기"
End Sub

'------------------------------------------------------------
' 피벗테이블 만들기 (Ctrl+Shift+J)
'   현재 영역을 원본으로 새 시트에 빈 피벗을 만든다.
'   필드는 필드 목록에서 직접 끌어 넣고, 정리는 Ctrl+Shift+T.
'------------------------------------------------------------
Public Sub 피벗만들기()
    Dim ws As Worksheet
    Dim 새시트 As Worksheet
    Dim 원본 As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim 이름 As String
    Dim i As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "피벗테이블 만들기"
        Exit Sub
    End If
    Set ws = ActiveSheet

    '--- 원본 범위 : 합계 행은 제외하고 잡는다 ---
    Set 원본 = 소계행제외(ActiveCell.CurrentRegion)
    If 원본 Is Nothing Then
        MsgBox "데이터를 찾지 못했습니다.", vbExclamation, "피벗테이블 만들기"
        Exit Sub
    End If
    If 원본.Rows.Count < 2 Then
        MsgBox "머리글과 데이터가 함께 있어야 합니다.", vbExclamation, "피벗테이블 만들기"
        Exit Sub
    End If

    On Error GoTo Fail

    '--- 새 시트 이름 정하기 (피벗1, 피벗2 ...) ---
    i = 1
    Do
        이름 = "피벗" & i
        If Not 시트있음(ws.Parent, 이름) Then Exit Do
        i = i + 1
    Loop While i < 1000

    Set 새시트 = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
    On Error Resume Next
    새시트.Name = 이름
    Err.Clear
    On Error GoTo Fail

    Set pc = ws.Parent.PivotCaches.Create(SourceType:=xlDatabase, _
                                          SourceData:=원본)
    Set pt = pc.CreatePivotTable(TableDestination:=새시트.Range("A3"))

    새시트.Activate
    새시트.Range("A3").Select

    ' 필드 목록 창 띄우기
    On Error Resume Next
    ActiveWorkbook.ShowPivotTableFieldList = True
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "피벗을 만들었습니다.  원본 " & ws.Name & "!" & _
                            원본.Address(False, False) & _
                            "   필드를 넣은 뒤 Ctrl+Shift+T 로 정리하세요."
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "피벗테이블을 만들지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "피벗테이블 만들기"
End Sub

Private Function 시트있음(ByVal wb As Workbook, ByVal 이름 As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(이름)
    Err.Clear
    On Error GoTo 0
    시트있음 = Not (sh Is Nothing)
End Function

'------------------------------------------------------------
' 설정 초기화
'   배포 전에 실행해서 공장 기본값으로 되돌린다.
'   자동 암호확인 = 켜짐,  기본 암호 = 1234
'------------------------------------------------------------
Public Sub 설정초기화()
    Dim 현재자동 As String
    Dim 현재암호 As String
    Dim shp As Object
    Dim ws As Worksheet
    Dim 지운수 As Long

    현재자동 = IIf(자동확인사용(), "켜짐", "꺼짐")
    현재암호 = 기본암호읽기()

    If MsgBox("설정을 기본값으로 되돌립니다." & vbCrLf & vbCrLf & _
              "   자동 암호확인 : " & 현재자동 & "  ->  켜짐" & vbCrLf & _
              "   기본 암호     : " & 현재암호 & "  ->  " & 기본암호_초기값 & vbCrLf & vbCrLf & _
              "열려 있는 파일의 단축키 목록(Ctrl+Shift+K)도 함께 지웁니다." & vbCrLf & vbCrLf & _
              "계속할까요?", vbOKCancel + vbQuestion, "설정 초기화") = vbCancel Then Exit Sub

    설정쓰기 "auto", "1"
    설정쓰기 "pwd", 기본암호_초기값

    ' 열려 있는 모든 시트의 치트시트 도형 제거
    On Error Resume Next
    For Each ws In Application.ActiveWorkbook.Worksheets
        Set shp = Nothing
        Set shp = ws.Shapes(치트시트이름)
        If Not shp Is Nothing Then
            shp.Delete
            지운수 = 지운수 + 1
        End If
    Next ws
    Err.Clear
    On Error GoTo 0

    MsgBox "설정을 기본값으로 되돌렸습니다." & vbCrLf & vbCrLf & _
           "   자동 암호확인 : 켜짐" & vbCrLf & _
           "   기본 암호     : " & 기본암호_초기값 & vbCrLf & _
           IIf(지운수 > 0, "   단축키 목록 " & 지운수 & "개 제거" & vbCrLf, "") & vbCrLf & _
           "배포용으로 복사하기 전에 Excel 을 완전히 종료하세요." & vbCrLf & _
           "복사 위치 : " & ThisWorkbook.FullName, _
           vbInformation, "설정 초기화"
End Sub

'------------------------------------------------------------
' 표 테두리 (Ctrl+Shift+U)
'   머리글 : 굵게 + 가운데 정렬 + 사각형 테두리
'   데이터 : 세로 구분선만 (행 사이 가로선 없음)
'   바깥 테두리와 맨 아래 선은 넣는다. 모두 같은 굵기.
'------------------------------------------------------------
Public Sub 표테두리()
    Dim ws As Worksheet
    Dim rng As Range
    Dim 머리글 As Range
    Dim 본문 As Range
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "표 테두리"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "표로 만들 범위를 선택한 뒤 실행해주세요.", vbExclamation, "표 테두리"
        Exit Sub
    End If

    If Selection.Cells.CountLarge >= 2 Then
        Set rng = Selection
    Else
        Set rng = 소계행제외(ActiveCell.CurrentRegion)
    End If

    If rng Is Nothing Then
        MsgBox "표로 만들 데이터를 찾지 못했습니다.", vbExclamation, "표 테두리"
        Exit Sub
    End If
    If rng.Areas.Count > 1 Then
        MsgBox "떨어져 있는 여러 범위는 한 번에 처리할 수 없습니다.", vbExclamation, "표 테두리"
        Exit Sub
    End If
    If rng.Rows.Count < 2 Then
        MsgBox "머리글과 데이터가 함께 있어야 합니다.", vbExclamation, "표 테두리"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    테두리보관 rng

    Set 머리글 = rng.Rows(1)
    Set 본문 = rng.Offset(1, 0).Resize(rng.Rows.Count - 1, rng.Columns.Count)

    '--- 기존 테두리 지우기 ---
    rng.Borders.LineStyle = xlNone

    '--- 바깥 테두리 ---
    선긋기 rng, xlEdgeTop
    선긋기 rng, xlEdgeBottom
    선긋기 rng, xlEdgeLeft
    선긋기 rng, xlEdgeRight

    '--- 세로 구분선 (전체) ---
    If rng.Columns.Count > 1 Then 선긋기 rng, xlInsideVertical

    '--- 머리글 아래 가로선 ---
    선긋기 머리글, xlEdgeBottom

    '--- 머리글 서식 ---
    머리글.Font.Bold = True
    머리글.HorizontalAlignment = xlCenter

    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    Application.OnUndo "표 테두리 취소", 접두어() & "표테두리되돌리기"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "표 테두리를 적용했습니다. (" & rng.Address(False, False) & ")" & _
                            "  |  Ctrl+Z 로 되돌리기"
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "표 테두리를 적용하지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "표 테두리"
End Sub

Private Sub 선긋기(ByVal rg As Range, ByVal 위치 As Long)
    On Error Resume Next
    With rg.Borders(위치)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Err.Clear
    On Error GoTo 0
End Sub

'--- 표 테두리 되돌리기용 보관 ---
Private Sub 테두리보관(ByVal rg As Range)
    Dim r As Long, c As Long, i As Long
    Dim nR As Long, nC As Long
    Dim vL() As Variant, vB() As Variant, vA() As Variant
    Dim 위치 As Variant

    mBdrOn = False
    If rg Is Nothing Then Exit Sub
    If rg.Cells.CountLarge > 20000 Then Exit Sub

    nR = rg.Rows.Count
    nC = rg.Columns.Count
    위치 = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)

    On Error GoTo Fail
    ReDim vL(1 To nR, 1 To nC, 0 To 3)
    ReDim vB(1 To nR, 1 To nC)
    ReDim vA(1 To nR, 1 To nC)

    For r = 1 To nR
        For c = 1 To nC
            For i = 0 To 3
                vL(r, c, i) = rg.Cells(r, c).Borders(위치(i)).LineStyle
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
' 표 테두리 되돌리기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 표테두리되돌리기()
    Dim rng As Range
    Dim r As Long, c As Long, i As Long
    Dim 위치 As Variant
    Dim oldScreen As Boolean

    If Not mBdrOn Or Len(mBdrAddr) = 0 Then
        MsgBox "되돌릴 표 테두리가 없습니다.", vbInformation, "표 테두리"
        Exit Sub
    End If

    위치 = Array(xlEdgeTop, xlEdgeBottom, xlEdgeLeft, xlEdgeRight)
    oldScreen = Application.ScreenUpdating

    On Error Resume Next
    Application.ScreenUpdating = False
    Set rng = Application.Workbooks(mBdrWb).Worksheets(mBdrWs).Range(mBdrAddr)

    rng.Borders.LineStyle = xlNone

    For r = 1 To rng.Rows.Count
        For c = 1 To rng.Columns.Count
            For i = 0 To 3
                If mBdrLine(r, c, i) <> xlNone Then
                    With rng.Cells(r, c).Borders(위치(i))
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

    Application.StatusBar = "표 테두리를 되돌렸습니다."
    상태표시줄예약
End Sub

'--- 셀 배경 지우기 되돌리기용 보관 ---
Private Sub 배경색보관(ByVal rg As Range)
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
' 셀 배경 지우기 되돌리기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 배경색되돌리기()
    Dim rng As Range
    Dim ar As Range
    Dim cel As Range
    Dim i As Long
    Dim oldScreen As Boolean

    If Not mBgOn Or Len(mBgAddr) = 0 Then
        MsgBox "되돌릴 배경색이 없습니다.", vbInformation, "셀 배경 지우기"
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

    Application.StatusBar = "셀 배경색을 되돌렸습니다."
    상태표시줄예약
End Sub
