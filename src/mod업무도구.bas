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
'    Ctrl+Shift+V  기준 열 중복 (색칠 · 그룹번호 · 추출)
'    Ctrl+Shift+C  셀 배경 지우기
'    Ctrl+Shift+U  표 테두리
'    Ctrl+Shift+A  필터 (자동필터 토글 / 고급필터 해제)
'    Ctrl+Shift+F  고급필터 실행
'    Ctrl+Shift+R  고급필터 조건 열 세팅
'    Ctrl+Shift+G  SUBTOTAL 합계
'    Ctrl+Alt+G    SUBTOTAL 개수 (채워진 칸 수)
'    Ctrl+Shift+E  시트 빼서 텍스트 저장
'    Ctrl+Shift+N  파일 경로 복사
'    Ctrl+Shift+Z  선택한 셀 값으로 필터
'    Ctrl+Shift+P  암호 설정
'    Ctrl+Shift+O  새 창 열기
'    Ctrl+Shift+K  단축키 목록 붙이기 / 지우기
'    Ctrl+Shift+L  다중 VLOOKUP (여러 건을 옆으로 펼치기)
'    Ctrl+Alt+F    값으로 찾기 (계산 결과 기준)
'    Ctrl+Shift+1  수식 아래로 채우기
'    Ctrl+Shift+2  인덱스 열 삽입
'    Ctrl+Shift+3  표 끝까지 선택 (빈 칸 무시)
'    Ctrl+Shift+4  위 칸 값으로 빈 칸 채우기
'    Ctrl+Shift+M  업무도구 메뉴
'
'    Ctrl+Alt+Z    가운데 정렬 / 기본 (토글)
'    Ctrl+Alt+S    선택 영역의 가운데로 (병합 없이)
'    Ctrl+Alt+A    빠른 채우기 (Ctrl+E 업그레이드, 표 끝까지)
'    Ctrl+Alt+Q    자동 채우기 (1,2,3 / 월,화,수 패턴, 표 끝까지)
'    Ctrl+Alt+E    머리글 위 행 제거
'    Ctrl+Alt+X    서식 복사 / 붙여넣기 (한 키로 번갈아)
'    Ctrl+Alt+D    두 열 대조 (양방향, 목록에 있는지)
'    Ctrl+Alt+B    열 머리글을 행 데이터로 (열 피벗 해제)
'    (자동)        저장된 파일을 닫을 때 열기암호 확인
'============================================================

Private Const 서식복사_유지초 As Long = 40
Public Const 도구버전 As String = "1.1.0"
Public Const 도구버전일자 As String = "2026-11-01"
Public Const 빌드시각 As String = "2026-11-01 18:50"
Public Const 제작자 As String = "SH"
Public Const 배포처 As String = "oditso.com"

Private Const 기본암호_초기값 As String = "1234"
Private Const 암호최대길이 As Long = 15
Private Const 메뉴이름 As String = "업무도구메뉴"
Private Const 치트시트이름 As String = "업무도구_단축키목록"
Private Const 기준열메뉴이름 As String = "기준열메뉴"
Private Const 다중조회메뉴이름 As String = "다중조회메뉴"
Private Const 두열대조메뉴이름 As String = "두열대조메뉴"
Private Const 두곳메뉴이름 As String = "두곳메뉴"
Private Const 쉼표메뉴이름 As String = "쉼표메뉴"
Private Const 열행메뉴이름 As String = "열행메뉴"

Private mBusy As Boolean               ' 재진입 방지
Private mSuppressOnce As Boolean       ' 지연 닫기 중 재질문 방지
Private mPendingCloseName As String    ' 지연 닫기 대상

' SUBTOTAL 합계 되돌리기용 보관
Private mUndoWb As String
Private mUndoWs As String
Private mUndoAddr As String
Private mUndoFormula As Variant
Private mUndoBold As Variant
Private mUndoNF As Variant
Private mUndoFunc As Long
Private mUndoTitle As String

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

' 합친 버튼이 호출할 때 각자의 완료 메시지를 막는다
Private mQuiet As Boolean
Private mDupGrp As Long, mDupHit As Long

' 기준 열 중복 찾기 - 고른 열 상태 보관
Private mColPick() As Boolean
Private mColWs As String
Private mColAddr As String

' 다중 VLOOKUP - 고른 열 상태
Private mLkKey() As Boolean
Private mLkVal() As Boolean
Private mLkWs As String
Private mLkAddr As String
Private mLkFormula As Boolean
Private mLkDedup As Boolean

' 열 머리글을 행 데이터로 - 고른 열 상태
'   mUpVal(i) = True  ->  그 열이 행으로 내려간다 (●)
'   False            ->  그대로 남는다 (○)
Private mUpVal() As Boolean
Private mUpWs As String
Private mUpAddr As String
Private mUpBlank As Boolean
Private mUpDown As Boolean
Private mUpName1 As String
Private mUpName2 As String

' 두 열 대조 - 고른 열 상태
Private mCmpA As Long
Private mCmpB As Long
Private mCmpWs As String
Private mCmpAddr As String
Private mCmpFormula As Boolean

' 머리글 위 행 제거 되돌리기용
Private mDelWb As String, mDelWs As String
Private mDelRows As Long, mDelCols As Long
Private mDelF As Variant, mDelN As Variant
Private mDelC As Variant, mDelI As Variant
Private mDelB As Variant, mDelA As Variant
Private mDelH As Variant
Private mDelOn As Boolean

' 서식 복사 보관
Private mFmtSrcWb As String, mFmtSrcWs As String, mFmtSrcAddr As String
Private mFmtSrcOn As Boolean
Private mFmtSrcTime As Date

' 같은 행끼리 비교 되돌리기용
Private mDifWb As String, mDifWs As String
Private mDifResAddr As String, mDifColAddr As String
Private mDifRes As Variant
Private mDifColor As Variant, mDifIndex As Variant
Private mDifOn As Boolean
Private mDifIns As Boolean
Private mDifInsCol As Long

' Ctrl 로 잡은 두 곳
Private mCmpDirWs As String
Private mCmpDirA As String, mCmpDirB As String

' 정렬 되돌리기용
Private mAlignWb As String, mAlignWs As String, mAlignAddr As String
Private mAlignVal As Variant, mAlignOn As Boolean

' 삽입 되돌리기용
Private mInsWb As String, mInsWs As String
Private mInsCol As Long, mInsCount As Long, mInsOn As Boolean

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
    Application.OnKey "^+v", p & "기준열중복"
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
    Application.OnKey "^+l", p & "다중조회"
    Application.OnKey "^%f", p & "값으로찾기"
    Application.OnKey "^+1", p & "수식채우기"
    Application.OnKey "^+2", p & "인덱스열"
    Application.OnKey "^+3", p & "끝까지선택"
    Application.OnKey "^+4", p & "위칸채우기"
    Application.OnKey "^+m", p & "매크로런처"

    Application.OnKey "^%z", p & "가운데정렬"
    Application.OnKey "^%s", p & "가운데걸치기"
    Application.OnKey "^%a", p & "빠른채우기"
    Application.OnKey "^%q", p & "자동채우기"
    Application.OnKey "^%e", p & "머리글위제거"
    Application.OnKey "^%x", p & "서식복사"
    Application.OnKey "^%d", p & "두열대조"
    Application.OnKey "^%b", p & "열행전환"
    Application.OnKey "^%g", p & "개수넣기"
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
    서식복사해제
    Application.OnKey "^%d"
    Application.OnKey "^%b"
    Application.OnKey "^%g"
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

    s = "Excel 업무도구 " & 도구버전 & "  (빌드 " & 빌드시각 & ")" & vbCrLf & _
        "Made by " & 제작자 & " in Korea  ·  " & 배포처 & vbCrLf & vbCrLf & _
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
          "[취소]    닫지 않고 그대로 둡니다.  (Esc)" & vbCrLf & vbCrLf & _
          "기본값 변경 : Ctrl+Shift+M → 끄기"

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

    Selection.NumberFormat = 쉼표형식문자열()

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
        MsgBox "피벗테이블 안의 셀을 선택한 뒤 실행해주세요.", vbExclamation, "클래식 피벗"
        Exit Sub
    End If

    On Error GoTo Fail
    Application.ScreenUpdating = False

    With pt
        .InGridDropZones = True                 ' 클래식 피벗테이블 레이아웃
        .RowAxisLayout xlTabularRow             ' 행 필드를 각각 별도 열로
    End With

    pt.ManualUpdate = True                      ' 바꿀 때마다 다시 계산하지 않도록

    For Each pf In pt.RowFields                 ' 모든 행 필드 부분합 제거
        For i = 1 To 12
            On Error Resume Next
            pf.Subtotals(i) = False
            Err.Clear
            On Error GoTo Fail
        Next i
    Next pf

    pt.RepeatAllLabels xlRepeatLabels           ' 모든 항목 레이블 반복

    On Error Resume Next
    pt.ManualUpdate = False                     ' 다시 자동 갱신으로
    Err.Clear
    On Error GoTo Fail

    Application.ScreenUpdating = oldScreen

    MsgBox "피벗 설정 완료!" & vbCrLf & _
           "· 클래식 피벗 표시" & vbCrLf & _
           "· 행 필드 열 분리" & vbCrLf & _
           "· 부분합 전부 제외" & vbCrLf & _
           "· 모든 항목 레이블 반복" & vbCrLf & vbCrLf & _
           "※ Ctrl+Z 로 되돌릴 수 없습니다.", vbInformation, "클래식 피벗"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not pt Is Nothing Then pt.ManualUpdate = False
    Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = oldScreen
    MsgBox "설정 중 오류가 발생했습니다." & vbCrLf & _
           "오류 번호: " & errNo & vbCrLf & errDesc, _
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
    Dim 셀색() As Variant
    Dim 셀색인덱스() As Variant
    Dim 보임() As Boolean
    Dim 보임범위 As Range
    Dim 영역 As Range
    Dim 시작 As Long
    Dim 보이는수 As Long
    Dim 색주소() As String
    Dim 색번호 As Long
    Dim i As Long

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
    ' Ctrl 로 여러 곳을 골라 잡았으면 셀 단위로 값끼리 비교한다
    If rng.Areas.Count > 1 Then
        중복값찾기 rng
        Exit Sub
    End If

    rowCnt = rng.Rows.Count
    colCnt = rng.Columns.Count

    If rowCnt < 2 Then
        MsgBox "2행 이상을 선택해주세요.", vbExclamation, "중복행 색칠"
        Exit Sub
    End If
    '--- 보이는 행 파악 (필터 / 숨긴 행 제외) ---
    ReDim 보임(1 To rowCnt)
    보이는수 = 0

    On Error Resume Next
    Set 보임범위 = rng.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo 0

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
        MsgBox "보이는 행이 2행 미만입니다." & vbCrLf & _
               "필터를 확인한 뒤 다시 실행해주세요.", vbExclamation, "중복행 색칠"
        Exit Sub
    End If

    '--- 실제로 검사할 행 수 기준으로 안내 ---
    If 보이는수 > 100000 Then
        If MsgBox(Format$(보이는수, "#,##0") & "행을 검사합니다. 시간이 걸릴 수 있습니다." & vbCrLf & _
                  IIf(보이는수 < rowCnt, _
                      "(전체 " & Format$(rowCnt, "#,##0") & "행 중 보이는 행만)" & vbCrLf, "") & _
                  "계속할까요?", vbOKCancel + vbQuestion, "중복행 색칠") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating

    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = rng.Value

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set 그룹색 = CreateObject("Scripting.Dictionary")
    그룹색.CompareMode = 1

    팔레트 = 색팔레트()

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
        ' 행 색을 한 번에 읽는다. 한 행 안에서 색이 섞여 있으면(Null)
        ' 그 행만 셀 단위로 따로 담는다.
        ReDim 이전색(1 To dupRows)
        ReDim 이전색인덱스(1 To dupRows)
        For r = 1 To dupRows
            With rng.Rows(대상행(r))
                이전색(r) = .Interior.Color
                이전색인덱스(r) = .Interior.ColorIndex
            End With
            If IsNull(이전색(r)) Or IsNull(이전색인덱스(r)) Then
                ReDim 셀색(1 To colCnt)
                ReDim 셀색인덱스(1 To colCnt)
                For c = 1 To colCnt
                    셀색(c) = rng.Cells(대상행(r), c).Interior.Color
                    셀색인덱스(c) = rng.Cells(대상행(r), c).Interior.ColorIndex
                Next c
                이전색(r) = 셀색
                이전색인덱스(r) = 셀색인덱스
            End If
        Next r
        mDupWb = rng.Worksheet.Parent.Name
        mDupWs = rng.Worksheet.Name
        mDupAddr = rng.Address
        mDupRows = 대상행
        mDupColor = 이전색
        mDupIndex = 이전색인덱스
        mDupCount = dupRows
    End If

    ' 3차 : 색깔별로 주소를 모아 한 번에 칠한다 (행마다 칠하면 매우 느림)
    dupRows = 0
    ReDim 색주소(0 To UBound(팔레트))
    For r = 1 To rowCnt
      If 보임(r) Then
        key = ""
        For c = 1 To colCnt
            key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            If Not 그룹색.Exists(key) Then
                그룹색.Add key, 그룹수 Mod (UBound(팔레트) + 1)
                그룹수 = 그룹수 + 1
            End If
            색번호 = 그룹색(key)
            If Len(색주소(색번호)) > 0 Then 색주소(색번호) = 색주소(색번호) & ","
            색주소(색번호) = 색주소(색번호) & rng.Rows(r).Address(False, False)
            dupRows = dupRows + 1

            ' 주소 문자열이 길어지면 그때그때 칠하고 비운다
            If Len(색주소(색번호)) > 200 Then
                rng.Worksheet.Range(색주소(색번호)).Interior.Color = 팔레트(색번호)
                색주소(색번호) = ""
            End If
        End If
      End If
    Next r

    For i = 0 To UBound(팔레트)
        If Len(색주소(i)) > 0 Then
            rng.Worksheet.Range(색주소(i)).Interior.Color = 팔레트(i)
        End If
    Next i

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
    메뉴항목 cb, "가운데 정렬 / 기본", "가운데정렬", "Ctrl+Alt+Z"
    메뉴항목 cb, "서식 복사 / 붙여넣기", "서식복사", "Ctrl+Alt+X"
    메뉴항목 cb, "선택 영역의 가운데로 (병합 없이)", "가운데걸치기", "Ctrl+Alt+S"
    메뉴항목 cb, "빠른 채우기 (Ctrl+E 업그레이드)", "빠른채우기", "Ctrl+Alt+A"
    메뉴항목 cb, "자동 채우기 (1,2,3 / 월,화,수 패턴)", "자동채우기", "Ctrl+Alt+Q"
    메뉴항목 cb, "위 칸 값으로 빈 칸 채우기", "위칸채우기", "Ctrl+Shift+4"
    메뉴항목 cb, "수식 아래로 채우기", "수식채우기", "Ctrl+Shift+1"
    메뉴항목 cb, "인덱스 열 삽입 (No)", "인덱스열", "Ctrl+Shift+2"
    메뉴항목 cb, "표 끝까지 선택", "끝까지선택", "Ctrl+Shift+3"
    메뉴항목 cb, "머리글 위 행 제거", "머리글위제거", "Ctrl+Alt+E"
    메뉴항목 cb, "행/열 자동맞춤", "행열자동맞춤", "Ctrl+Shift+W"
    메뉴항목 cb, "첫 행 고정 / 해제", "첫행고정", "Ctrl+Shift+S"
    메뉴항목 cb, "시트 일괄정리 (모든 시트 A1으로)", "시트일괄정리", "Ctrl+Shift+H"

    메뉴항목 cb, "피벗테이블 만들기 (새 시트)", "피벗만들기", "Ctrl+Shift+J", True
    메뉴항목 cb, "클래식 피벗 정리 (되돌리기 불가)", "클래식피벗", "Ctrl+Shift+T"
    메뉴항목 cb, "중복행 색칠", "중복행색칠", "Ctrl+Shift+D"
    메뉴항목 cb, "기준 열 중복 (색칠 · 그룹번호 · 추출)", "기준열중복", "Ctrl+Shift+V"
    메뉴항목 cb, "두 열 비교 (목록끼리 대조 · 같은 행끼리)", "두열대조", "Ctrl+Alt+D"
    메뉴항목 cb, "다중 VLOOKUP (그룹화하여 펼치기)", "다중조회", "Ctrl+Shift+L"
    메뉴항목 cb, "열 머리글을 행 데이터로 (열 피벗 해제)", "열행전환", "Ctrl+Alt+B"
    메뉴항목 cb, "셀 배경 지우기 (범위 선택 필요)", "중복색지우기", "Ctrl+Shift+C"

    메뉴항목 cb, "필터 켜기 / 끄기", "자동필터", "Ctrl+Shift+A", True
    메뉴항목 cb, "선택한 셀 값으로 필터", "선택값으로필터", "Ctrl+Shift+Z"
    메뉴항목 cb, "고급필터 실행", "고급필터", "Ctrl+Shift+F"
    메뉴항목 cb, "고급필터 조건 열 세팅", "조건블록만들기", "Ctrl+Shift+R"
    메뉴항목 cb, "고급필터 범위 확인", "고급필터범위확인", ""
    메뉴항목 cb, "SUBTOTAL 합계", "소계넣기", "Ctrl+Shift+G"
    메뉴항목 cb, "SUBTOTAL 개수 (채워진 칸 수)", "개수넣기", "Ctrl+Alt+G"

    메뉴항목 cb, "시트 빼서 텍스트 저장 (txt→dat 변환용)", "시트빼서텍스트저장", "Ctrl+Shift+E", True
    메뉴항목 cb, "파일 경로 복사", "파일경로복사", "Ctrl+Shift+N"
    메뉴항목 cb, "값으로 찾기 (계산 결과 기준)", "값으로찾기", "Ctrl+Alt+F"

    메뉴항목 cb, "암호 설정", "암호설정", "Ctrl+Shift+P", True
    메뉴항목 cb, 자동확인캡션(), "자동확인전환", ""
    메뉴항목 cb, "기본 암호 변경 (현재 : " & 기본암호읽기() & ")", "기본암호변경", ""

    메뉴항목 cb, "단축키 목록 붙이기 / 지우기", "치트시트", "Ctrl+Shift+K", True
    메뉴항목 cb, "쉼표 형식 모양 :  " & 쉼표형식한줄(), "쉼표형식메뉴", ""

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
    Dim 건너뜀 As Long
    Dim 남은행 As Long

    Set 소계행제외 = rg
    If rg Is Nothing Then Exit Function
    If rg.Rows.Count < 3 Then Exit Function

    On Error GoTo Done

    ' 머리글이 아닌 행을 최대 3줄까지 위에서 걷어낸다
    건너뜀 = 0
    Do While 건너뜀 < 3
        남은행 = rg.Rows.Count - 건너뜀
        If 남은행 < 3 Then Exit Do
        If Not 머리글행아님(rg, 건너뜀 + 1) Then Exit Do
        건너뜀 = 건너뜀 + 1
    Loop

    If 건너뜀 > 0 Then
        Set 소계행제외 = rg.Offset(건너뜀, 0).Resize(rg.Rows.Count - 건너뜀, rg.Columns.Count)
    End If

Done:
End Function

'------------------------------------------------------------
' 그 행이 머리글로 보기 어려운지 판정
'   1) SUBTOTAL 수식이 있다
'   2) 값이 있는 칸이 대부분 숫자다  (머리글은 보통 글자)
'   3) 값이 있는 칸이 아랫행보다 훨씬 적다
'------------------------------------------------------------
Private Function 머리글행아님(ByVal rg As Range, ByVal 행번호 As Long) As Boolean
    Dim i As Long
    Dim 값있음 As Long, 숫자 As Long
    Dim 아래값 As Long
    Dim v As Variant
    Dim f As String
    Dim nC As Long

    On Error GoTo Done
    nC = rg.Columns.Count
    If 행번호 + 1 > rg.Rows.Count Then Exit Function

    For i = 1 To nC
        ' 1) SUBTOTAL 수식
        If rg.Cells(행번호, i).HasFormula Then
            f = UCase$(CStr(rg.Cells(행번호, i).Formula))
            If InStr(f, "SUBTOTAL") > 0 Then
                머리글행아님 = True
                Exit Function
            End If
        End If

        v = rg.Cells(행번호, i).Value
        If Len(CStr(v)) > 0 Then
            값있음 = 값있음 + 1
            ' 셀이 실제로 숫자로 저장돼 있을 때만 센다.
            '   IsDate 를 쓰면 안 된다 - "2025년1월" 같은 머리글 글자를
            '   날짜로 해석해버려서 진짜 머리글 행을 걷어내 버린다.
            '   날짜도 세지 않는다 - 월별 데이터는 머리글이 진짜 날짜 셀인
            '   경우가 있고, 소계 행은 날짜가 아니라 숫자로 채워진다.
            Select Case VarType(v)
                Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
                    숫자 = 숫자 + 1
            End Select
        End If

        If Len(CStr(rg.Cells(행번호 + 1, i).Value)) > 0 Then 아래값 = 아래값 + 1
    Next i

    If 값있음 = 0 Then
        머리글행아님 = True                      ' 완전히 빈 행
        Exit Function
    End If

    ' 2) 값이 있는 칸의 절반 이상이 숫자/날짜면 머리글로 보기 어렵다
    If 숫자 * 2 >= 값있음 Then
        머리글행아님 = True
        Exit Function
    End If

    ' 3) 아랫행에 견줘 채워진 칸이 절반도 안 되면 머리글이 아니다
    If 아래값 >= 4 Then
        If 값있음 * 2 < 아래값 Then 머리글행아님 = True
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
    소계핵심 9, "SUBTOTAL 합계", False, "Ctrl+Shift+G"
End Sub

Public Sub 개수넣기()
    소계핵심 3, "SUBTOTAL 개수", True, "Ctrl+Alt+G"
End Sub

'------------------------------------------------------------
' 합계(9) / 개수(3) 공통 본체
'   함수번호 : SUBTOTAL 의 첫 인수.  9 = SUM, 3 = COUNTA
'   서식강제 : True 면 #,##0 으로 바꾼다 (통화 서식 열에 개수가 들어갈 때)
'------------------------------------------------------------
Private Sub 소계핵심(ByVal 함수번호 As Long, _
                     ByVal 제목 As String, _
                     ByVal 서식강제 As Boolean, _
                     ByVal 키표시 As String)
    Dim ws As Worksheet
    Dim sel As Range
    Dim c As Range
    Dim 보임 As Range, 영역 As Range
    Dim 시작행 As Long, 끝행 As Long
    Dim 넣은수 As Long
    Dim oldScreen As Boolean
    Dim 마지막 As Long
    Dim i As Long
    Dim 전체행 As Long, 보이는행 As Long
    Dim 덧붙임 As String
    Dim 이전수식() As Variant
    Dim 이전굵게() As Variant
    Dim 이전서식() As Variant

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, 제목
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "넣을 셀을 선택한 뒤 실행해주세요.", vbExclamation, 제목
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "떨어져 있는 여러 범위는 한 번에 처리할 수 없습니다.", vbExclamation, 제목
        Exit Sub
    End If
    If sel.Rows.Count > 1 Then
        MsgBox "한 행만 선택해주세요." & vbCrLf & _
               "넣을 칸들을 가로로 선택하면 각 열마다 넣어드립니다.", _
               vbExclamation, 제목
        Exit Sub
    End If

    ' 머리글 한 줄 건너뛴 다음 행부터
    시작행 = sel.Row + 2
    If 시작행 > ws.Rows.Count Then
        MsgBox "데이터를 찾을 수 없습니다.", vbExclamation, 제목
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
        MsgBox "계산할 데이터가 없습니다." & vbCrLf & vbCrLf & _
               "넣을 칸에 커서를 두고 실행하세요." & vbCrLf & _
               "바로 아래 칸은 머리글로 보고 건너뜁니다.", vbExclamation, 제목
        Exit Sub
    End If

    '--- 마지막으로 같은 키를 누른 자리와 같으면 원래대로 되돌린다 ---
    If mUndoOn And Len(mUndoAddr) > 0 And mUndoFunc = 함수번호 Then
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
    ReDim 이전서식(1 To sel.Cells.Count)
    i = 0
    For Each c In sel.Cells
        i = i + 1
        이전수식(i) = c.Formula
        이전굵게(i) = c.Font.Bold
        이전서식(i) = c.NumberFormat
    Next c
    mUndoWb = ws.Parent.Name
    mUndoWs = ws.Name
    mUndoAddr = sel.Address
    mUndoFormula = 이전수식
    mUndoBold = 이전굵게
    mUndoNF = 이전서식
    mUndoFunc = 함수번호
    mUndoTitle = 제목
    mUndoOn = True

    For Each c In sel.Cells
        c.Formula = "=SUBTOTAL(" & 함수번호 & "," & _
                    ws.Cells(시작행, c.Column).Address(False, False) & ":" & _
                    ws.Cells(끝행, c.Column).Address(False, False) & ")"
        c.Font.Bold = True
        If 서식강제 Then c.NumberFormat = "#,##0"
        넣은수 = 넣은수 + 1
    Next c

    '--- 개수일 때는 전체 / 보이는 행 수도 같이 알려준다 ---
    덧붙임 = ""
    If 함수번호 = 3 Then
        전체행 = 끝행 - 시작행 + 1
        보이는행 = 전체행
        On Error Resume Next
        If ws.FilterMode Then
            Set 보임 = ws.Range(ws.Cells(시작행, sel.Cells(1, 1).Column), _
                                ws.Cells(끝행, sel.Cells(1, 1).Column)) _
                         .SpecialCells(xlCellTypeVisible)
            If Not 보임 Is Nothing Then
                보이는행 = 0
                For Each 영역 In 보임.Areas
                    보이는행 = 보이는행 + 영역.Rows.Count
                Next 영역
            End If
        End If
        Err.Clear
        On Error GoTo Fail
        덧붙임 = "  |  전체 " & Format$(전체행, "#,##0") & "행 / 보이는 " & _
                 Format$(보이는행, "#,##0") & "행"
    End If

    Application.ScreenUpdating = oldScreen

    ' Ctrl+Z 로 되돌릴 수 있게 등록
    On Error Resume Next
    Application.OnUndo 제목 & " 취소", 접두어() & "소계되돌리기"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = 제목 & " " & 넣은수 & "칸을 넣었습니다. (" & _
                            시작행 & "행 ~ " & 끝행 & "행)" & 덧붙임 & "  |  " & _
                            "Ctrl+Z 또는 " & 키표시 & " 로 되돌리기"
    상태표시줄예약
    Exit Sub

Fail:
    Application.ScreenUpdating = oldScreen
    MsgBox "넣지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, 제목
End Sub

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

'------------------------------------------------------------
' 되돌리기 - SUBTOTAL 합계 / 개수 공용
'------------------------------------------------------------
Public Sub 소계되돌리기()
    Dim ws As Worksheet
    Dim rg As Range
    Dim i As Long
    Dim oldScreen As Boolean
    Dim 제목 As String
    Dim errNo As Long, errDesc As String

    제목 = mUndoTitle
    If Len(제목) = 0 Then 제목 = "SUBTOTAL 합계"

    If Len(mUndoAddr) = 0 Then
        MsgBox "되돌릴 " & 제목 & "가 없습니다.", vbInformation, 제목
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

    Application.StatusBar = 제목 & "를 되돌렸습니다."
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "되돌리지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, 제목
End Sub

'------------------------------------------------------------
' 중복행 색칠 되돌리기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 중복색되돌리기()
    Dim ws As Worksheet
    Dim rng As Range
    Dim r As Long, c As Long
    Dim 색모음 As Object
    Dim 빈주소 As String
    Dim 키색 As String
    Dim 키 As Variant
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
    Set 색모음 = CreateObject("Scripting.Dictionary")

    ' 색이 같은 행끼리 주소를 모아 한 번에 되돌린다
    For r = 1 To mDupCount
        If IsArray(mDupColor(r)) Then
            ' 한 행 안에서 색이 섞여 있던 경우만 셀 단위
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
            If Len(빈주소) > 0 Then 빈주소 = 빈주소 & ","
            빈주소 = 빈주소 & rng.Rows(mDupRows(r)).Address(False, False)
            If Len(빈주소) > 200 Then
                rng.Worksheet.Range(빈주소).Interior.ColorIndex = xlColorIndexNone
                빈주소 = ""
            End If
        Else
            키색 = CStr(mDupColor(r))
            If 색모음.Exists(키색) Then
                색모음(키색) = 색모음(키색) & "," & rng.Rows(mDupRows(r)).Address(False, False)
            Else
                색모음.Add 키색, rng.Rows(mDupRows(r)).Address(False, False)
            End If
            If Len(색모음(키색)) > 200 Then
                rng.Worksheet.Range(색모음(키색)).Interior.Color = CLng(키색)
                색모음(키색) = ""
            End If
        End If
    Next r

    If Len(빈주소) > 0 Then rng.Worksheet.Range(빈주소).Interior.ColorIndex = xlColorIndexNone
    For Each 키 In 색모음.Keys
        If Len(색모음(키)) > 0 Then
            rng.Worksheet.Range(색모음(키)).Interior.Color = CLng(키)
        End If
    Next 키

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
    ElseIf VarType(v) = vbDate Then
        ' 셀이 진짜 날짜로 저장돼 있을 때만 날짜 범위로 건다.
        '   IsDate 를 쓰면 안 된다 - "2023-06-13" 같은 글자도 True 라서,
        '   텍스트로 들어온 날짜 열에 숫자(일련번호) 조건이 걸려 0건이 된다.
        crit = Format$(v, "yyyy-mm-dd")
        fr.AutoFilter Field:=fld, _
                      Criteria1:=">=" & CDbl(v), _
                      Operator:=xlAnd, _
                      Criteria2:="<" & (CDbl(v) + 1)
    ElseIf IsNumeric(v) And VarType(v) <> vbString Then
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

    '--- 3차 시도 : 보이는 값 그대로 목록으로 거른다 ---
    '    Criteria1 문자열은 Excel 이 날짜로 해석해버릴 수 있다.
    '    xlFilterValues 는 그 해석을 거치지 않아 텍스트 날짜에도 걸린다.
    If 남은행 = 0 And Len(c.Text) > 0 Then
        On Error Resume Next
        fr.AutoFilter Field:=fld, Criteria1:=Array(c.Text), Operator:=xlFilterValues
        Err.Clear
        On Error GoTo Fail
        남은행 = 표시된행수(fr)
        If 남은행 > 0 Then crit = c.Text
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
    s = s & "Ctrl+Shift+X    쉼표 형식  (음수·0 모양은 Ctrl+Shift+M 에서 설정)" & vbLf
    s = s & "Ctrl+Shift+Y    텍스트 나누기 - 일반으로  (텍스트 숫자를 숫자로)" & vbLf
    s = s & "Ctrl+Shift+I    텍스트 나누기 - 텍스트로" & vbLf
    s = s & "Ctrl+Shift+U    표 테두리  (머리글 굵게 + 세로 구분선)" & vbLf
    s = s & "Ctrl+Alt+Z      가운데 정렬 / 기본  (토글)" & vbLf
    s = s & "Ctrl+Alt+X      서식 복사 / 붙여넣기  (한 키로 번갈아)" & vbLf
    s = s & "Ctrl+Alt+S      선택 영역의 가운데로  (병합 없이 걸쳐 보이기)" & vbLf
    s = s & "Ctrl+Shift+W    행/열 자동맞춤" & vbLf
    s = s & "Ctrl+Shift+S    첫 행 고정 / 해제  (머리글 자동 인식)" & vbLf
    s = s & "Ctrl+Shift+H    시트 일괄정리  (모든 시트 A1으로, 확대 100%)" & vbLf & vbLf
    s = s & "[ 채우기 · 선택 ]" & vbLf
    s = s & "Ctrl+Alt+A      빠른 채우기  (Ctrl+E 업그레이드, 범위 지정 없이 표 끝까지)" & vbLf
    s = s & "Ctrl+Alt+Q      자동 채우기  (1,2,3 / 월,화,수 패턴을 표 끝까지)" & vbLf
    s = s & "Ctrl+Shift+1    수식 아래로 채우기  (표 끝까지)" & vbLf
    s = s & "Ctrl+Shift+4    위 칸 값으로 빈 칸 채우기  (병합 푼 표 정리)" & vbLf
    s = s & "Ctrl+Shift+2    인덱스 열 삽입  (No, =ROW()-머리글행)" & vbLf
    s = s & "Ctrl+Shift+3    표 끝까지 선택  (빈 칸 무시)" & vbLf
    s = s & "Ctrl+Alt+E      머리글 위 행 제거  (내보내기 전 정리)" & vbLf
    s = s & vbLf
    s = s & "[ 데이터 ]" & vbLf
    s = s & "Ctrl+Shift+J    피벗테이블 만들기  (새 시트, 범위 자동)" & vbLf
    s = s & "Ctrl+Shift+T    클래식 피벗 정리  (부분합 제거, 되돌리기 불가)" & vbLf
    s = s & "Ctrl+Shift+D    중복행 색칠  (Ctrl로 골라 잡으면 값끼리 비교)" & vbLf
    s = s & "Ctrl+Shift+V    기준 열 중복  (열 골라서 색칠 · 그룹번호 · 새 시트로 추출)" & vbLf
    s = s & "Ctrl+Shift+L    다중 VLOOKUP  (같은 값끼리 묶어 펼치기 · 그룹화)" & vbLf
    s = s & "Ctrl+Alt+D      두 열 비교  (목록끼리 대조 · 같은 행끼리 비교)" & vbLf
    s = s & "Ctrl+Alt+B      열 머리글을 행 데이터로  (1월 2월 3월 열 -> 한 열로)" & vbLf
    s = s & "Ctrl+Shift+C    셀 배경 지우기  (선택 범위의 채우기 색 제거)" & vbLf
    s = s & "Ctrl+Shift+G    SUBTOTAL 합계  (머리글 위 칸에서, 다시 누르면 취소)" & vbLf
    s = s & "Ctrl+Alt+G      SUBTOTAL 개수  (채워진 칸 수, 필터 따라감)" & vbLf & vbLf
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
    s = s & "Ctrl+Alt+F      값으로 찾기  (수식 결과 기준, 찾은 칸 모두 선택)" & vbLf
    s = s & "Ctrl+Shift+P    암호 설정  (열기 암호, 기본값 1234)" & vbLf
    s = s & "Ctrl+Shift+K    이 목록 붙이기 / 지우기" & vbLf & vbLf
    s = s & "저장된 파일을 닫을 때 열기 암호가 없으면 자동으로 묻습니다." & vbLf & vbLf
    s = s & "Made by " & 제작자 & " in Korea  ·  " & 배포처
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
    '   [일반] 로 되돌릴 때는 셀 서식이 [텍스트](@) 로 남아 있으면
    '   값만 숫자가 되고 리본에는 계속 [텍스트] 로 보인다.
    '   그래서 먼저 셀 서식을 General 로 풀어 준다.
    If 서식 = xlGeneralFormat Then
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
                          FieldInfo:=Array(1, 서식), _
                          TrailingMinusNumbers:=True
    Next col

    ' Excel 은 마지막 [텍스트 나누기] 설정을 기억했다가
    ' 이후 붙여넣기에도 그대로 적용한다.
    ' 구분 기호를 꺼 둔 채로 두면 탭으로 나뉜 값을 붙여넣어도
    ' 한 칸에 다 들어가므로, 기본값(탭 구분)으로 되돌려 놓는다.
    텍스트나누기설정복원

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
    텍스트나누기설정복원
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
    Dim oldScreen As Boolean, oldEvents As Boolean
    Dim oldCalc As Long
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

    ' 화면·재계산·이벤트를 잠근다 (안 잠그면 큰 표에서 매우 느림)
    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    oldEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    '--- 새 시트 이름 정하기 : 피벗_원본시트명 ---
    이름 = 피벗시트이름(ws.Parent, ws.Name)

    ' 원본 시트 바로 앞에 만든다
    Set 새시트 = ws.Parent.Worksheets.Add(Before:=ws)
    On Error Resume Next
    새시트.Name = 이름
    Err.Clear
    On Error GoTo Fail

    Set pc = ws.Parent.PivotCaches.Create(SourceType:=xlDatabase, _
                                          SourceData:=원본)
    Set pt = pc.CreatePivotTable(TableDestination:=새시트.Range("B3"))

    ' 원본 위치를 A1 에 적어 둔다
    On Error Resume Next
    With 새시트.Range("A1")
        .Value = "원본 : " & ws.Name & "!" & 원본.Address(False, False)
        .Font.Size = 9
        .Font.Color = RGB(110, 110, 115)
    End With
    새시트.Columns("A").ColumnWidth = 3
    Err.Clear
    On Error GoTo Fail

    새시트.Activate
    새시트.Range("B3").Select

    ' 잠근 것들을 되돌린 뒤 필드 목록 창을 띄운다
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
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
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
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
              "   기본 암호     : " & 현재암호 & "  ->  " & 기본암호_초기값 & vbCrLf & _
              "   쉼표 형식     : " & 쉼표형식한줄() & vbCrLf & _
              "                  ->  3,333 / △3,333 / -" & vbCrLf & vbCrLf & _
              "열려 있는 파일의 단축키 목록(Ctrl+Shift+K)도 함께 지웁니다." & vbCrLf & vbCrLf & _
              "계속할까요?", vbOKCancel + vbQuestion, "설정 초기화") = vbCancel Then Exit Sub

    설정쓰기 "auto", "1"
    설정쓰기 "pwd", 기본암호_초기값
    설정쓰기 "cmaColor", "R"
    설정쓰기 "cmaNeg", "T"
    설정쓰기 "cmaZero", "D"
    설정쓰기 "cmaPlus", "N"
    설정쓰기 "cmaPColor", "K"

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
           "   쉼표 형식     : 3,333 / △3,333 / -" & vbCrLf & _
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

'--- 피벗 시트 이름 만들기 : 피벗_원본명 (31자 제한, 중복 회피) ---
Private Function 피벗시트이름(ByVal wb As Workbook, ByVal 원본명 As String) As String
    Dim 기본 As String
    Dim 후보 As String
    Dim 꼬리 As String
    Dim i As Long

    ' 시트 이름에 쓸 수 없는 문자 정리
    원본명 = Replace(원본명, ":", "")
    원본명 = Replace(원본명, "\", "")
    원본명 = Replace(원본명, "/", "")
    원본명 = Replace(원본명, "?", "")
    원본명 = Replace(원본명, "*", "")
    원본명 = Replace(원본명, "[", "")
    원본명 = Replace(원본명, "]", "")

    기본 = "피벗_" & 원본명
    If Len(기본) > 31 Then 기본 = Left$(기본, 31)

    If Not 시트있음(wb, 기본) Then
        피벗시트이름 = 기본
        Exit Function
    End If

    ' 이미 있으면 뒤에 번호를 붙이되 31자를 넘지 않게 자른다
    For i = 2 To 999
        꼬리 = CStr(i)
        후보 = 기본
        If Len(후보) + Len(꼬리) > 31 Then
            후보 = Left$(후보, 31 - Len(꼬리))
        End If
        후보 = 후보 & 꼬리
        If Not 시트있음(wb, 후보) Then
            피벗시트이름 = 후보
            Exit Function
        End If
    Next i

    피벗시트이름 = "피벗_" & Format$(Now, "hhnnss")
End Function

'------------------------------------------------------------
' 중복 값 찾기 (Ctrl+Shift+D 에서 여러 곳을 골라 잡았을 때)
'   행 단위가 아니라 골라 잡은 셀들의 값끼리 비교해 같은 것끼리 같은 색.
'------------------------------------------------------------
Private Sub 중복값찾기(ByVal rng As Range)
    Dim dic As Object
    Dim 그룹색 As Object
    Dim 팔레트 As Variant
    Dim ar As Range, cel As Range
    Dim key As String
    Dim 그룹수 As Long
    Dim 검사수 As Long, 색칠수 As Long
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If rng Is Nothing Then Exit Sub

    검사수 = 0
    On Error Resume Next
    검사수 = rng.Cells.CountLarge
    Err.Clear
    On Error GoTo 0

    If 검사수 < 2 Then
        MsgBox "2칸 이상 골라 주세요.", vbExclamation, "중복 값 찾기"
        Exit Sub
    End If
    If 검사수 > 50000 Then
        MsgBox "골라 잡은 칸이 너무 많습니다. (" & Format$(검사수, "#,##0") & "칸)" & vbCrLf & _
               "5만 칸 이하로 줄여주세요.", vbExclamation, "중복 값 찾기"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set 그룹색 = CreateObject("Scripting.Dictionary")
    그룹색.CompareMode = 1

    팔레트 = 색팔레트()

    ' 1차 : 값별 개수
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

    ' 되돌리기용 보관
    배경색보관 rng

    ' 2차 : 2개 이상인 값에 그룹별 색
    For Each ar In rng.Areas
        For Each cel In ar.Cells
            key = CStr(cel.Value)
            If Len(key) > 0 Then
                If dic(key) >= 2 Then
                    If Not 그룹색.Exists(key) Then
                        그룹색.Add key, 팔레트(그룹수 Mod (UBound(팔레트) + 1))
                        그룹수 = 그룹수 + 1
                    End If
                    cel.Interior.Color = 그룹색(key)
                    색칠수 = 색칠수 + 1
                End If
            End If
        Next cel
    Next ar

    Application.ScreenUpdating = oldScreen

    If 색칠수 > 0 And mBgOn Then
        On Error Resume Next
        Application.OnUndo "중복 값 찾기 취소", 접두어() & "배경색되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    If 색칠수 = 0 Then
        MsgBox "골라 잡은 " & Format$(검사수, "#,##0") & "칸 중 같은 값이 없습니다.", _
               vbInformation, "중복 값 찾기"
    Else
        MsgBox "골라 잡은 " & Format$(검사수, "#,##0") & "칸을 값끼리 비교했습니다." & vbCrLf & _
               "같은 값 " & 그룹수 & "종류 / " & Format$(색칠수, "#,##0") & "칸을 표시했습니다." & _
               IIf(그룹수 > UBound(팔레트) + 1, vbCrLf & "종류가 " & (UBound(팔레트) + 1) & _
                   "가지를 넘어 일부 색이 반복됩니다.", "") & vbCrLf & vbCrLf & _
               "Ctrl+Z 로 되돌릴 수 있습니다." & vbCrLf & _
               "색만 지우려면 : 범위 선택 후 Ctrl+Shift+C", _
               vbInformation, "중복 값 찾기"
    End If
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "중복 값을 찾지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "중복 값 찾기"
End Sub

'============================================================
'  기준 열 중복 찾기 (Ctrl+Shift+V)
'    비교할 열을 골라서, 그 열들의 값 조합이 같은 행을 색칠한다.
'    [중복된 항목 제거] 와 비슷하지만 지우지 않고 표시만 한다.
'============================================================

Public Sub 기준열중복()
    Dim ws As Worksheet
    Dim 데이터 As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "기준 열 중복 찾기"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set 데이터 = 기준열범위(ws)
    If 데이터 Is Nothing Then Exit Sub

    ' 다른 표로 옮겨갔으면 고른 열을 초기화
    If StrComp(mColWs, ws.Name, vbTextCompare) <> 0 Or _
       StrComp(mColAddr, 데이터.Address, vbTextCompare) <> 0 Then
        ReDim mColPick(1 To 데이터.Columns.Count)
        mColWs = ws.Name
        mColAddr = 데이터.Address
    End If

    On Error Resume Next
    If UBound(mColPick) <> 데이터.Columns.Count Then
        ReDim mColPick(1 To 데이터.Columns.Count)
    End If
    Err.Clear
    On Error GoTo 0

    기준열메뉴 데이터
End Sub

'--- 대상 표 범위 결정 ---
Private Function 기준열범위(ByVal ws As Worksheet) As Range
    Dim rg As Range

    If TypeName(Selection) <> "Range" Then
        MsgBox "표 안의 셀을 선택한 뒤 실행해주세요.", vbExclamation, "기준 열 중복 찾기"
        Exit Function
    End If

    If Selection.Cells.CountLarge >= 2 And Selection.Areas.Count = 1 Then
        Set rg = Selection
    Else
        Set rg = 소계행제외(ActiveCell.CurrentRegion)
    End If

    If rg Is Nothing Then
        MsgBox "표를 찾지 못했습니다.", vbExclamation, "기준 열 중복 찾기"
        Exit Function
    End If
    If rg.Rows.Count < 3 Then
        MsgBox "머리글 1행 + 데이터 2행 이상이 필요합니다.", vbExclamation, "기준 열 중복 찾기"
        Exit Function
    End If
    If rg.Columns.Count > 60 Then
        MsgBox "열이 " & rg.Columns.Count & "개로 너무 많습니다." & vbCrLf & _
               "비교할 열이 있는 범위만 선택해주세요.", vbExclamation, "기준 열 중복 찾기"
        Exit Function
    End If

    Set 기준열범위 = rg
End Function

'--- 열 고르기 팝업 메뉴 ---
Private Sub 기준열메뉴(ByVal 데이터 As Range)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim i As Long
    Dim 고른수 As Long
    Dim 이름 As String

    On Error Resume Next
    Application.CommandBars(기준열메뉴이름).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=기준열메뉴이름, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, _
                                         Temporary:=True)

    For i = 1 To 데이터.Columns.Count
        If mColPick(i) Then 고른수 = 고른수 + 1
    Next i

    메뉴제목 cb, "기준 열 중복 찾기        (" & 데이터.Address(False, False) & ")"
    메뉴제목 cb, "■  찍은 열들의 값이 모두 같은 행을 찾습니다."

    For i = 1 To 데이터.Columns.Count
        이름 = Trim$(CStr(데이터.Cells(1, i).Value))
        If Len(이름) = 0 Then 이름 = "(빈 머리글) " & Split(데이터.Cells(1, i).Address(True, False), "$")(0)

        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mColPick(i), "■  ", "□  ") & 이름
        btn.OnAction = 접두어() & "기준열토글" & Format$(i, "00")
        btn.Style = msoButtonCaption
        If i = 1 Then btn.BeginGroup = True
    Next i

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  중복 색칠 + 그룹번호   (한 번에)"
    btn.OnAction = 접두어() & "중복색번호넣기"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True
    btn.Enabled = (고른수 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  중복 색칠   (" & 고른수 & "개 선택됨)"
    btn.OnAction = 접두어() & "기준열실행"
    btn.Style = msoButtonCaption
    btn.Enabled = (고른수 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  중복 그룹번호 넣기   (표 오른쪽에 열 2개)"
    btn.OnAction = 접두어() & "중복번호넣기"
    btn.Style = msoButtonCaption
    btn.Enabled = (고른수 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  중복인 행만 새 시트로   (모든 열)"
    btn.OnAction = 접두어() & "중복행추출"
    btn.Style = msoButtonCaption
    btn.Enabled = (고른수 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  고른 열 고유 목록 새 시트로   (중복 제거)"
    btn.OnAction = 접두어() & "고유목록추출"
    btn.Style = msoButtonCaption
    btn.Enabled = (고른수 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "모두 해제"
    btn.OnAction = 접두어() & "기준열모두해제"
    btn.Style = msoButtonCaption

    cb.ShowPopup
    Exit Sub

Fail:
    MsgBox "메뉴를 표시하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "기준 열 중복 찾기"
End Sub

Public Sub 기준열모두해제()
    Dim i As Long
    On Error Resume Next
    For i = LBound(mColPick) To UBound(mColPick)
        mColPick(i) = False
    Next i
    Err.Clear
    On Error GoTo 0
    기준열중복
End Sub

'--- 열 하나를 켜고 끈 뒤 메뉴를 다시 연다 ---
Private Sub 기준열체크(ByVal 번호 As Long)
    On Error Resume Next
    mColPick(번호) = Not mColPick(번호)
    Err.Clear
    On Error GoTo 0
    기준열중복
End Sub

'--- 실제 중복 찾기 ---
Public Sub 기준열실행()
    Dim ws As Worksheet
    Dim 데이터 As Range
    Dim dic As Object, 그룹색 As Object
    Dim 팔레트 As Variant
    Dim arr As Variant
    Dim 보임범위 As Range, 영역 As Range
    Dim 보임() As Boolean
    Dim 대상행() As Variant
    Dim 이전색() As Variant, 이전색인덱스() As Variant
    Dim 셀색() As Variant, 셀색인덱스() As Variant
    Dim key As String
    Dim r As Long, c As Long, i As Long, 시작 As Long
    Dim rowCnt As Long, colCnt As Long
    Dim 보이는수 As Long
    Dim 색주소() As String
    Dim 색번호 As Long
    Dim 그룹수 As Long, dupRows As Long
    Dim 고른열 As String
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = ActiveSheet
    Set 데이터 = ws.Range(mColAddr)
    rowCnt = 데이터.Rows.Count
    colCnt = 데이터.Columns.Count

    '--- 보이는 행 파악 (머리글 행 제외) ---
    ReDim 보임(1 To rowCnt)
    On Error Resume Next
    Set 보임범위 = 데이터.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo 0

    If 보임범위 Is Nothing Then
        For r = 2 To rowCnt
            보임(r) = True
            보이는수 = 보이는수 + 1
        Next r
    Else
        For Each 영역 In 보임범위.Areas
            시작 = 영역.Row - 데이터.Row + 1
            For r = 시작 To 시작 + 영역.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not 보임(r) Then
                        보임(r) = True
                        보이는수 = 보이는수 + 1
                    End If
                End If
            Next r
        Next 영역
    End If

    If 보이는수 < 2 Then
        MsgBox "보이는 데이터가 2행 미만입니다.", vbExclamation, "기준 열 중복 찾기"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = 데이터.Value

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set 그룹색 = CreateObject("Scripting.Dictionary")
    그룹색.CompareMode = 1
    팔레트 = 색팔레트()

    ' 1차 : 고른 열 조합으로 개수 세기
    For r = 2 To rowCnt
      If 보임(r) Then
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

    ' 2차 : 대상 행 수집
    ReDim 대상행(1 To rowCnt)
    For r = 2 To rowCnt
      If 보임(r) Then
        key = ""
        For c = 1 To colCnt
            If mColPick(c) Then key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            dupRows = dupRows + 1
            대상행(dupRows) = r
        End If
      End If
    Next r

    ' 되돌리기용 보관
    If dupRows > 0 Then
        ' 행 색을 한 번에 읽는다. 섞여 있는(Null) 행만 셀 단위로.
        ReDim 이전색(1 To dupRows)
        ReDim 이전색인덱스(1 To dupRows)
        For r = 1 To dupRows
            With 데이터.Rows(대상행(r))
                이전색(r) = .Interior.Color
                이전색인덱스(r) = .Interior.ColorIndex
            End With
            If IsNull(이전색(r)) Or IsNull(이전색인덱스(r)) Then
                ReDim 셀색(1 To colCnt)
                ReDim 셀색인덱스(1 To colCnt)
                For c = 1 To colCnt
                    셀색(c) = 데이터.Cells(대상행(r), c).Interior.Color
                    셀색인덱스(c) = 데이터.Cells(대상행(r), c).Interior.ColorIndex
                Next c
                이전색(r) = 셀색
                이전색인덱스(r) = 셀색인덱스
            End If
        Next r
        mDupWb = ws.Parent.Name
        mDupWs = ws.Name
        mDupAddr = 데이터.Address
        mDupRows = 대상행
        mDupColor = 이전색
        mDupIndex = 이전색인덱스
        mDupCount = dupRows
    End If

    ' 3차 : 색깔별로 주소를 모아 한 번에 칠한다
    dupRows = 0
    ReDim 색주소(0 To UBound(팔레트))
    For r = 2 To rowCnt
      If 보임(r) Then
        key = ""
        For c = 1 To colCnt
            If mColPick(c) Then key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            If Not 그룹색.Exists(key) Then
                그룹색.Add key, 그룹수 Mod (UBound(팔레트) + 1)
                그룹수 = 그룹수 + 1
            End If
            색번호 = 그룹색(key)
            If Len(색주소(색번호)) > 0 Then 색주소(색번호) = 색주소(색번호) & ","
            색주소(색번호) = 색주소(색번호) & 데이터.Rows(r).Address(False, False)
            dupRows = dupRows + 1
            If Len(색주소(색번호)) > 200 Then
                ws.Range(색주소(색번호)).Interior.Color = 팔레트(색번호)
                색주소(색번호) = ""
            End If
        End If
      End If
    Next r

    For i = 0 To UBound(팔레트)
        If Len(색주소(i)) > 0 Then ws.Range(색주소(i)).Interior.Color = 팔레트(i)
    Next i

    Application.ScreenUpdating = oldScreen

    If dupRows > 0 Then
        On Error Resume Next
        Application.OnUndo "기준 열 중복 찾기 취소", 접두어() & "중복색되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    ' 고른 열 이름 모으기
    For c = 1 To colCnt
        If mColPick(c) Then
            If Len(고른열) > 0 Then 고른열 = 고른열 & " + "
            고른열 = 고른열 & Trim$(CStr(데이터.Cells(1, c).Value))
        End If
    Next c

    If mQuiet Then Exit Sub

    If dupRows = 0 Then
        MsgBox "중복된 행이 없습니다." & vbCrLf & vbCrLf & _
               "기준 열 : " & 고른열 & vbCrLf & _
               "검사 범위 : " & 데이터.Address(False, False) & _
               IIf(보이는수 < rowCnt - 1, "  (보이는 " & 보이는수 & "행만)", ""), _
               vbInformation, "기준 열 중복 찾기"
    Else
        MsgBox "중복 " & 그룹수 & "종류 / " & Format$(dupRows, "#,##0") & "행을 표시했습니다." & vbCrLf & vbCrLf & _
               "기준 열 : " & 고른열 & vbCrLf & _
               "검사 범위 : " & 데이터.Address(False, False) & _
               IIf(보이는수 < rowCnt - 1, "  (보이는 " & 보이는수 & "행만)", "") & vbCrLf & vbCrLf & _
               "Ctrl+Z 로 되돌릴 수 있습니다." & vbCrLf & _
               "색만 지우려면 : 범위 선택 후 Ctrl+Shift+C", _
               vbInformation, "기준 열 중복 찾기"
    End If
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "중복을 찾지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "기준 열 중복 찾기"
End Sub

'------------------------------------------------------------
' 중복 그룹번호 / 개수 열 넣기  (Ctrl+Shift+V 팝업에서)
'   고른 열 조합이 2번 이상 나오는 행에만 번호와 개수를 넣는다.
'   번호는 처음 나온 순서대로 1, 2, 3...  중복 아닌 행은 빈칸.
'   표 오른쪽에 열 두 개를 삽입하고, Ctrl+Z 로 그 열이 지워진다.
'------------------------------------------------------------
Public Sub 중복번호넣기()
    Dim ws As Worksheet
    Dim 데이터 As Range
    Dim dic As Object, 그룹번호 As Object
    Dim arr As Variant
    Dim 보임범위 As Range, 영역 As Range
    Dim 보임() As Boolean
    Dim 출력() As Variant
    Dim key As String
    Dim r As Long, c As Long, 시작 As Long
    Dim rowCnt As Long, colCnt As Long
    Dim 보이는수 As Long
    Dim 그룹수 As Long, dupRows As Long
    Dim 고른수 As Long
    Dim 오른쪽열 As Long, 머리글행 As Long
    Dim 고른열 As String
    Dim oldScreen As Boolean, oldCalc As Long
    Dim 잠갔음 As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Len(mColAddr) = 0 Then
        MsgBox "먼저 Ctrl+Shift+V 를 눌러 기준 열을 골라주세요.", _
               vbInformation, "중복 그룹번호"
        Exit Sub
    End If

    Set ws = ActiveSheet
    Set 데이터 = ws.Range(mColAddr)
    rowCnt = 데이터.Rows.Count
    colCnt = 데이터.Columns.Count

    For c = 1 To colCnt
        If mColPick(c) Then 고른수 = 고른수 + 1
    Next c
    If 고른수 < 1 Then
        MsgBox "기준 열을 하나 이상 골라주세요.", vbExclamation, "중복 그룹번호"
        Exit Sub
    End If

    '--- 보이는 행 파악 (머리글 행 제외) ---
    ReDim 보임(1 To rowCnt)
    On Error Resume Next
    Set 보임범위 = 데이터.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo Fail

    If 보임범위 Is Nothing Then
        For r = 2 To rowCnt
            보임(r) = True
            보이는수 = 보이는수 + 1
        Next r
    Else
        For Each 영역 In 보임범위.Areas
            시작 = 영역.Row - 데이터.Row + 1
            For r = 시작 To 시작 + 영역.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not 보임(r) Then
                        보임(r) = True
                        보이는수 = 보이는수 + 1
                    End If
                End If
            Next r
        Next 영역
    End If

    If 보이는수 < 2 Then
        MsgBox "보이는 데이터가 2행 미만입니다.", vbExclamation, "중복 그룹번호"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    잠갔음 = True
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    arr = 데이터.Value

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set 그룹번호 = CreateObject("Scripting.Dictionary")
    그룹번호.CompareMode = 1

    ' 1차 : 고른 열 조합으로 개수 세기
    For r = 2 To rowCnt
      If 보임(r) Then
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

    ' 2차 : 중복인 것만 처음 나온 순서대로 번호를 주고 출력 배열을 채운다
    ReDim 출력(1 To rowCnt - 1, 1 To 2)
    For r = 2 To rowCnt
      If 보임(r) Then
        key = ""
        For c = 1 To colCnt
            If mColPick(c) Then key = key & CStr(arr(r, c)) & Chr$(1)
        Next c
        If dic(key) >= 2 Then
            If Not 그룹번호.Exists(key) Then
                그룹수 = 그룹수 + 1
                그룹번호.Add key, 그룹수
            End If
            출력(r - 1, 1) = 그룹번호(key)
            출력(r - 1, 2) = dic(key)
            dupRows = dupRows + 1
        End If
      End If
    Next r

    If dupRows = 0 Then
        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreen
        잠갔음 = False
        For c = 1 To colCnt
            If mColPick(c) Then
                If Len(고른열) > 0 Then 고른열 = 고른열 & " + "
                고른열 = 고른열 & Trim$(CStr(데이터.Cells(1, c).Value))
            End If
        Next c
        If Not mQuiet Then
            MsgBox "중복된 행이 없습니다." & vbCrLf & vbCrLf & _
                   "기준 열 : " & 고른열 & vbCrLf & _
                   "검사 범위 : " & 데이터.Address(False, False), _
                   vbInformation, "중복 그룹번호"
        End If
        Exit Sub
    End If

    '--- 표 오른쪽에 열 두 개 삽입 ---
    머리글행 = 데이터.Row
    오른쪽열 = 데이터.Column + colCnt
    ws.Range(ws.Columns(오른쪽열), ws.Columns(오른쪽열 + 1)).Insert Shift:=xlToRight

    mInsWb = ws.Parent.Name
    mInsWs = ws.Name
    mInsCol = 오른쪽열
    mInsCount = 2
    mInsOn = True

    ' 삽입한 열은 왼쪽 서식을 물려받으므로 일반으로 되돌린다
    With ws.Range(ws.Cells(머리글행, 오른쪽열), ws.Cells(머리글행 + rowCnt - 1, 오른쪽열 + 1))
        .ClearFormats
        .NumberFormat = "General"
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(머리글행, 오른쪽열), ws.Cells(머리글행, 오른쪽열 + 1))
        .Value = Array("중복그룹", "중복개수")
        .Font.Bold = True
    End With

    ws.Cells(머리글행 + 1, 오른쪽열).Resize(rowCnt - 1, 2).Value = 출력

    ws.Range(ws.Columns(오른쪽열), ws.Columns(오른쪽열 + 1)).AutoFit

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    잠갔음 = False

    On Error Resume Next
    Application.OnUndo "중복 그룹번호 취소", 접두어() & "삽입열되돌리기"
    Err.Clear
    On Error GoTo Fail

    For c = 1 To colCnt
        If mColPick(c) Then
            If Len(고른열) > 0 Then 고른열 = 고른열 & " + "
            고른열 = 고른열 & Trim$(CStr(데이터.Cells(1, c).Value))
        End If
    Next c

    mDupGrp = 그룹수
    mDupHit = dupRows
    If mQuiet Then Exit Sub

    MsgBox "중복 " & Format$(그룹수, "#,##0") & "종류 / " & _
           Format$(dupRows, "#,##0") & "행에 번호를 넣었습니다." & vbCrLf & vbCrLf & _
           "기준 열 : " & 고른열 & vbCrLf & _
           "넣은 열 : " & 열문자(오른쪽열) & "(중복그룹) · " & _
                          열문자(오른쪽열 + 1) & "(중복개수)" & vbCrLf & _
           "검사 범위 : " & 데이터.Address(False, False) & _
           IIf(보이는수 < rowCnt - 1, "  (보이는 " & Format$(보이는수, "#,##0") & "행만)", "") & vbCrLf & vbCrLf & _
           "중복 아닌 행은 빈칸입니다." & vbCrLf & _
           "Ctrl+Z 로 넣은 열 두 개를 지울 수 있습니다.", _
           vbInformation, "중복 그룹번호"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    If 잠갔음 Then
        Application.Calculation = oldCalc
        Application.ScreenUpdating = oldScreen
    End If
    MsgBox "번호를 넣지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "중복 그룹번호"
End Sub

'------------------------------------------------------------
' 중복 색칠 + 그룹번호 한 번에  (Ctrl+Shift+V 팝업에서)
'   기존 두 기능을 순서대로 불러 쓴다. 코드를 복사하지 않는다.
'   색은 mDup* 에, 삽입한 열은 mIns* 에 따로 보관되므로 서로 안 덮어쓴다.
'   OnUndo 는 마지막에 등록한 것만 남으므로, 둘을 부른 뒤
'   합친 되돌리기로 한 번 더 등록해 Ctrl+Z 한 번에 둘 다 되돌린다.
'------------------------------------------------------------
Public Sub 중복색번호넣기()
    Dim 고른수 As Long
    Dim c As Long

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Len(mColAddr) = 0 Then
        MsgBox "먼저 Ctrl+Shift+V 를 눌러 기준 열을 골라주세요.", _
               vbInformation, "중복 색칠 + 그룹번호"
        Exit Sub
    End If

    On Error Resume Next
    For c = LBound(mColPick) To UBound(mColPick)
        If mColPick(c) Then 고른수 = 고른수 + 1
    Next c
    Err.Clear
    On Error GoTo 0

    If 고른수 < 1 Then
        MsgBox "기준 열을 하나 이상 골라주세요.", vbExclamation, "중복 색칠 + 그룹번호"
        Exit Sub
    End If

    ' 이번 호출로 열이 들어갔는지 판정하기 위해 먼저 내린다
    mInsOn = False
    mDupGrp = 0
    mDupHit = 0

    On Error GoTo Fail
    mQuiet = True
    기준열실행
    중복번호넣기
    mQuiet = False
    On Error GoTo 0

    If Not mInsOn Then
        MsgBox "중복된 행이 없습니다.", vbInformation, "중복 색칠 + 그룹번호"
        Exit Sub
    End If

    On Error Resume Next
    Application.OnUndo "중복 색칠 + 그룹번호 취소", 접두어() & "중복색번호되돌리기"
    Err.Clear
    On Error GoTo 0

    MsgBox "중복 " & Format$(mDupGrp, "#,##0") & "종류 / " & _
           Format$(mDupHit, "#,##0") & "행을 칠하고 번호를 넣었습니다." & vbCrLf & vbCrLf & _
           "중복그룹 번호와 색이 짝을 이룹니다." & vbCrLf & _
           "묶음이 많아 색이 다시 돌면 번호로 구분하세요." & vbCrLf & vbCrLf & _
           "Ctrl+Z 로 색과 열 두 개가 함께 되돌아갑니다.", _
           vbInformation, "중복 색칠 + 그룹번호"
    Exit Sub

Fail:
    mQuiet = False
    MsgBox "처리하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, _
           vbExclamation, "중복 색칠 + 그룹번호"
End Sub

'--- 되돌리기 : 열 두 개를 지우고 색을 복원한다 ---
Public Sub 중복색번호되돌리기()
    On Error Resume Next
    삽입열되돌리기
    중복색되돌리기
    Err.Clear
    On Error GoTo 0
    Application.StatusBar = "중복 색칠과 그룹번호를 되돌렸습니다."
    상태표시줄예약
End Sub

'--- 열 토글 래퍼 (CommandBar OnAction 은 인수를 못 넘겨 번호별로 만든다) ---
Public Sub 기준열토글01()
    기준열체크 1
End Sub
Public Sub 기준열토글02()
    기준열체크 2
End Sub
Public Sub 기준열토글03()
    기준열체크 3
End Sub
Public Sub 기준열토글04()
    기준열체크 4
End Sub
Public Sub 기준열토글05()
    기준열체크 5
End Sub
Public Sub 기준열토글06()
    기준열체크 6
End Sub
Public Sub 기준열토글07()
    기준열체크 7
End Sub
Public Sub 기준열토글08()
    기준열체크 8
End Sub
Public Sub 기준열토글09()
    기준열체크 9
End Sub
Public Sub 기준열토글10()
    기준열체크 10
End Sub
Public Sub 기준열토글11()
    기준열체크 11
End Sub
Public Sub 기준열토글12()
    기준열체크 12
End Sub
Public Sub 기준열토글13()
    기준열체크 13
End Sub
Public Sub 기준열토글14()
    기준열체크 14
End Sub
Public Sub 기준열토글15()
    기준열체크 15
End Sub
Public Sub 기준열토글16()
    기준열체크 16
End Sub
Public Sub 기준열토글17()
    기준열체크 17
End Sub
Public Sub 기준열토글18()
    기준열체크 18
End Sub
Public Sub 기준열토글19()
    기준열체크 19
End Sub
Public Sub 기준열토글20()
    기준열체크 20
End Sub
Public Sub 기준열토글21()
    기준열체크 21
End Sub
Public Sub 기준열토글22()
    기준열체크 22
End Sub
Public Sub 기준열토글23()
    기준열체크 23
End Sub
Public Sub 기준열토글24()
    기준열체크 24
End Sub
Public Sub 기준열토글25()
    기준열체크 25
End Sub
Public Sub 기준열토글26()
    기준열체크 26
End Sub
Public Sub 기준열토글27()
    기준열체크 27
End Sub
Public Sub 기준열토글28()
    기준열체크 28
End Sub
Public Sub 기준열토글29()
    기준열체크 29
End Sub
Public Sub 기준열토글30()
    기준열체크 30
End Sub
Public Sub 기준열토글31()
    기준열체크 31
End Sub
Public Sub 기준열토글32()
    기준열체크 32
End Sub
Public Sub 기준열토글33()
    기준열체크 33
End Sub
Public Sub 기준열토글34()
    기준열체크 34
End Sub
Public Sub 기준열토글35()
    기준열체크 35
End Sub
Public Sub 기준열토글36()
    기준열체크 36
End Sub
Public Sub 기준열토글37()
    기준열체크 37
End Sub
Public Sub 기준열토글38()
    기준열체크 38
End Sub
Public Sub 기준열토글39()
    기준열체크 39
End Sub
Public Sub 기준열토글40()
    기준열체크 40
End Sub
Public Sub 기준열토글41()
    기준열체크 41
End Sub
Public Sub 기준열토글42()
    기준열체크 42
End Sub
Public Sub 기준열토글43()
    기준열체크 43
End Sub
Public Sub 기준열토글44()
    기준열체크 44
End Sub
Public Sub 기준열토글45()
    기준열체크 45
End Sub
Public Sub 기준열토글46()
    기준열체크 46
End Sub
Public Sub 기준열토글47()
    기준열체크 47
End Sub
Public Sub 기준열토글48()
    기준열체크 48
End Sub
Public Sub 기준열토글49()
    기준열체크 49
End Sub
Public Sub 기준열토글50()
    기준열체크 50
End Sub
Public Sub 기준열토글51()
    기준열체크 51
End Sub
Public Sub 기준열토글52()
    기준열체크 52
End Sub
Public Sub 기준열토글53()
    기준열체크 53
End Sub
Public Sub 기준열토글54()
    기준열체크 54
End Sub
Public Sub 기준열토글55()
    기준열체크 55
End Sub
Public Sub 기준열토글56()
    기준열체크 56
End Sub
Public Sub 기준열토글57()
    기준열체크 57
End Sub
Public Sub 기준열토글58()
    기준열체크 58
End Sub
Public Sub 기준열토글59()
    기준열체크 59
End Sub
Public Sub 기준열토글60()
    기준열체크 60
End Sub

'--- 중복 표시용 20색 팔레트 (밝은 10 + 한 단계 진한 10) ---
Private Function 색팔레트() As Variant
    색팔레트 = Array( _
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
'  수식 아래로 채우기 (Ctrl+Shift+1)
'============================================================
Public Sub 수식채우기()
    Dim ws As Worksheet
    Dim sel As Range
    Dim blk As Range
    Dim c As Range
    Dim 끝행 As Long
    Dim 채운수 As Long
    Dim 덮어쓸행 As Long
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "수식 아래로 채우기"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "수식이 있는 칸을 선택한 뒤 실행해주세요.", vbExclamation, "수식 아래로 채우기"
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "떨어져 있는 여러 범위는 한 번에 처리할 수 없습니다.", _
               vbExclamation, "수식 아래로 채우기"
        Exit Sub
    End If
    If sel.Rows.Count > 1 Then
        MsgBox "한 행만 선택해주세요." & vbCrLf & _
               "여러 열을 가로로 선택하면 각 열마다 아래로 채웁니다.", _
               vbExclamation, "수식 아래로 채우기"
        Exit Sub
    End If

    '--- 표 끝 행 찾기 ---
    Set blk = 소계행제외(sel.Cells(1, 1).CurrentRegion)
    If blk Is Nothing Then
        MsgBox "표를 찾지 못했습니다.", vbExclamation, "수식 아래로 채우기"
        Exit Sub
    End If
    끝행 = blk.Row + blk.Rows.Count - 1

    If 끝행 <= sel.Row Then
        MsgBox "아래로 채울 행이 없습니다." & vbCrLf & _
               "표 마지막 행에 있거나 표를 찾지 못했습니다.", _
               vbExclamation, "수식 아래로 채우기"
        Exit Sub
    End If

    '--- 아래에 이미 값이 있으면 확인 ---
    덮어쓸행 = 0
    For Each c In sel.Cells
        덮어쓸행 = 덮어쓸행 + Application.CountA(ws.Range(ws.Cells(sel.Row + 1, c.Column), _
                                                          ws.Cells(끝행, c.Column)))
    Next c

    If 덮어쓸행 > 0 Then
        If MsgBox("아래쪽에 이미 값이 " & Format$(덮어쓸행, "#,##0") & "칸 있습니다." & vbCrLf & _
                  "덮어쓸까요?" & vbCrLf & vbCrLf & _
                  "채울 범위 : " & (sel.Row + 1) & "행 ~ " & 끝행 & "행" & vbCrLf & _
                  "Ctrl+Z 로 되돌릴 수 있습니다.", _
                  vbOKCancel + vbExclamation, "수식 아래로 채우기") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    ' 되돌리기용 보관 (채울 자리 전체)
    Dim 되돌 As Range
    Set 되돌 = ws.Range(ws.Cells(sel.Row, sel.Column), _
                        ws.Cells(끝행, sel.Column + sel.Columns.Count - 1))
    조건열보관 되돌

    ' 서식은 건드리지 않고 수식/값만 채운다 (FillDown 은 서식까지 복사)
    For Each c In sel.Cells
        ws.Range(ws.Cells(sel.Row, c.Column), ws.Cells(끝행, c.Column)).Formula = c.Formula
        채운수 = 채운수 + 1
    Next c

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "수식 채우기 취소", 접두어() & "조건열되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "열 " & 채운수 & "개를 " & (sel.Row + 1) & "행 ~ " & 끝행 & "행까지 채웠습니다." & _
                            IIf(mCondOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "채우지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "수식 아래로 채우기"
End Sub

'============================================================
'  인덱스 열 삽입 (Ctrl+Shift+2)
'============================================================
Public Sub 인덱스열()
    Dim ws As Worksheet
    Dim blk As Range
    Dim 머리글행 As Long, 끝행 As Long, 왼쪽열 As Long
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "인덱스 열 삽입"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then Exit Sub

    Set blk = 소계행제외(ActiveCell.CurrentRegion)
    If blk Is Nothing Then
        MsgBox "표를 찾지 못했습니다.", vbExclamation, "인덱스 열 삽입"
        Exit Sub
    End If
    If blk.Rows.Count < 2 Then
        MsgBox "머리글과 데이터가 함께 있어야 합니다.", vbExclamation, "인덱스 열 삽입"
        Exit Sub
    End If

    머리글행 = blk.Row
    끝행 = blk.Row + blk.Rows.Count - 1
    왼쪽열 = blk.Column

    If MsgBox("표 맨 왼쪽에 인덱스 열을 넣습니다." & vbCrLf & vbCrLf & _
              "위치 : " & Split(ws.Cells(1, 왼쪽열).Address(True, False), "$")(0) & "열 앞에 삽입" & vbCrLf & _
              "머리글 : No" & vbCrLf & _
              "수식 : =ROW()-" & 머리글행 & vbCrLf & _
              "범위 : " & (머리글행 + 1) & "행 ~ " & 끝행 & "행" & vbCrLf & vbCrLf & _
              "Ctrl+Z 로 되돌릴 수 있습니다.", _
              vbOKCancel + vbQuestion, "인덱스 열 삽입") = vbCancel Then Exit Sub

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    ws.Columns(왼쪽열).Insert Shift:=xlToRight

    mInsWb = ws.Parent.Name
    mInsWs = ws.Name
    mInsCol = 왼쪽열
    mInsCount = 1
    mInsOn = True

    With ws.Cells(머리글행, 왼쪽열)
        .Value = "No"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(머리글행 + 1, 왼쪽열), ws.Cells(끝행, 왼쪽열))
        .Formula = "=ROW()-" & 머리글행
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(왼쪽열).AutoFit

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    Application.OnUndo "인덱스 열 취소", 접두어() & "삽입열되돌리기"
    Err.Clear
    On Error GoTo Fail

    Application.StatusBar = "인덱스 열을 넣었습니다. (" & (끝행 - 머리글행) & "행)" & _
                            "  |  Ctrl+Z 로 되돌리기"
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "인덱스 열을 넣지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "인덱스 열 삽입"
End Sub

'--- 삽입한 열 되돌리기 ---
Public Sub 삽입열되돌리기()
    Dim ws As Worksheet
    If Not mInsOn Or mInsCount < 1 Then
        MsgBox "되돌릴 삽입 열이 없습니다.", vbInformation, "업무도구"
        Exit Sub
    End If
    On Error Resume Next
    Set ws = Application.Workbooks(mInsWb).Worksheets(mInsWs)
    ws.Range(ws.Columns(mInsCol), ws.Columns(mInsCol + mInsCount - 1)).Delete Shift:=xlToLeft
    mInsOn = False
    mInsCount = 0
    Err.Clear
    On Error GoTo 0
    Application.StatusBar = "삽입한 열을 되돌렸습니다."
    상태표시줄예약
End Sub

'============================================================
'  값으로 찾기 (Ctrl+Alt+F)
'    Ctrl+F 는 기본이 [수식]에서 찾아 계산 결과가 안 잡힌다.
'    이 도구는 [값] 기준으로 찾아 해당 칸을 모두 선택한다.
'============================================================
Public Sub 값으로찾기()
    Dim ws As Worksheet
    Dim 범위 As Range
    Dim 찾은 As Range
    Dim c As Range
    Dim 첫주소 As String
    Dim 찾을값 As String
    Dim 개수 As Long
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "값으로 찾기"
        Exit Sub
    End If
    Set ws = ActiveSheet

    ' 2칸 이상 선택했으면 그 안에서, 아니면 시트 전체에서
    If TypeName(Selection) = "Range" Then
        If Selection.Cells.CountLarge >= 2 Then
            Set 범위 = Selection
        Else
            Set 범위 = ws.UsedRange
        End If
    Else
        Set 범위 = ws.UsedRange
    End If

    찾을값 = InputBox("찾을 값  (계산 결과 기준, 찾은 칸 모두 선택)" & vbCrLf & _
                      "찾은 뒤 Ctrl+Shift+D 로 값 중복 확인" & vbCrLf & vbCrLf & _
                      "범위 : " & 범위.Address(False, False), _
                      "값으로 찾기")

    If StrPtr(찾을값) = 0 Then Exit Sub
    If Len(찾을값) = 0 Then Exit Sub

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    Set c = 범위.Find(What:=찾을값, LookIn:=xlValues, LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False)

    If c Is Nothing Then
        Application.ScreenUpdating = oldScreen
        MsgBox "'" & 찾을값 & "' 을(를) 찾지 못했습니다." & vbCrLf & vbCrLf & _
               "찾는 범위 : " & 범위.Address(False, False), _
               vbInformation, "값으로 찾기"
        Exit Sub
    End If

    첫주소 = c.Address
    Do
        If 찾은 Is Nothing Then
            Set 찾은 = c
        Else
            Set 찾은 = Application.Union(찾은, c)
        End If
        개수 = 개수 + 1
        If 개수 >= 1000 Then Exit Do
        Set c = 범위.FindNext(c)
        If c Is Nothing Then Exit Do
    Loop While c.Address <> 첫주소

    찾은.Select

    Application.ScreenUpdating = oldScreen

    Application.StatusBar = "'" & 찾을값 & "' " & Format$(개수, "#,##0") & "칸을 선택했습니다." & _
                            IIf(개수 >= 1000, "  (1,000칸까지만)", "") & _
                            "   |  Ctrl+Shift+D 로 값 중복 확인"
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "찾지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "값으로 찾기"
End Sub

'============================================================
'  다중 VLOOKUP (Ctrl+Shift+L)
'    VLOOKUP 은 첫 번째 일치만 가져온다.
'    이 도구는 일치하는 여러 건을 옆으로 펼쳐 준다.
'============================================================
Public Sub 다중조회()
    Dim ws As Worksheet
    Dim 데이터 As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "다중 VLOOKUP"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set 데이터 = 기준열범위(ws)
    If 데이터 Is Nothing Then Exit Sub

    If StrComp(mLkWs, ws.Name, vbTextCompare) <> 0 Or _
       StrComp(mLkAddr, 데이터.Address, vbTextCompare) <> 0 Then
        ReDim mLkKey(1 To 데이터.Columns.Count)
        ReDim mLkVal(1 To 데이터.Columns.Count)
        mLkWs = ws.Name
        mLkAddr = 데이터.Address
        mLkFormula = False
        mLkDedup = True                       ' 기본 : 같은 조합은 한 번씩만
    End If

    On Error Resume Next
    If UBound(mLkKey) <> 데이터.Columns.Count Then
        ReDim mLkKey(1 To 데이터.Columns.Count)
        ReDim mLkVal(1 To 데이터.Columns.Count)
    End If
    Err.Clear
    On Error GoTo 0

    다중조회메뉴 데이터
End Sub

Private Sub 다중조회메뉴(ByVal 데이터 As Range)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim i As Long
    Dim 기준수 As Long, 값수 As Long
    Dim 이름 As String

    On Error Resume Next
    Application.CommandBars(다중조회메뉴이름).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=다중조회메뉴이름, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    For i = 1 To 데이터.Columns.Count
        If mLkKey(i) Then 기준수 = 기준수 + 1
        If mLkVal(i) Then 값수 = 값수 + 1
    Next i

    메뉴제목 cb, "다중 VLOOKUP  -  같은 값끼리 묶어 펼치기 (그룹화)"

    '--- ① 묶을 기준 ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "①  묶을 기준  (여러 개 가능)"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    For i = 1 To 데이터.Columns.Count
        이름 = 열이름(데이터, i)
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mLkKey(i), "     ■  ", "     □  ") & 이름
        btn.OnAction = 접두어() & "조회기준토글" & Format$(i, "00")
        btn.Style = msoButtonCaption
    Next i

    '--- ② 가져올 값 ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "②  가져올 값  (여러 개 가능)"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    For i = 1 To 데이터.Columns.Count
        이름 = 열이름(데이터, i)
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mLkVal(i), "     ●  ", "     ○  ") & 이름
        btn.OnAction = 접두어() & "조회값토글" & Format$(i, "00")
        btn.Style = msoButtonCaption
    Next i

    '--- ③ 실행 ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  고유 목록으로 새 시트   (기준마다 한 줄)"
    btn.OnAction = 접두어() & "다중조회새시트"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True
    btn.Enabled = (기준수 > 0 And 값수 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  원본 옆에 펼치기   (행 수 그대로)"
    btn.OnAction = 접두어() & "다중조회실행"
    btn.Style = msoButtonCaption
    btn.Enabled = (기준수 > 0 And 값수 > 0)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     같은 값 : " & IIf(mLkDedup, "한 번씩만  ->  중복 포함", "중복 포함  ->  한 번씩만")
    btn.OnAction = 접두어() & "조회중복전환"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     옆에 펼칠 때 : " & IIf(mLkFormula, "수식으로 (자동 갱신)", "값으로 (빠름)")
    btn.OnAction = 접두어() & "조회방식전환"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     모두 해제      (기준 " & 기준수 & " / 값 " & 값수 & ")"
    btn.OnAction = 접두어() & "조회모두해제"
    btn.Style = msoButtonCaption

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "메뉴를 표시하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "다중 VLOOKUP"
End Sub

Private Function 열이름(ByVal rg As Range, ByVal i As Long) As String
    Dim s As String
    On Error Resume Next
    s = Trim$(CStr(rg.Cells(1, i).Value))
    Err.Clear
    On Error GoTo 0
    If Len(s) = 0 Then s = "(빈 머리글) " & Split(rg.Cells(1, i).Address(True, False), "$")(1)
    열이름 = s
End Function

Public Sub 조회방식전환()
    mLkFormula = Not mLkFormula
    다중조회
End Sub

Public Sub 조회중복전환()
    mLkDedup = Not mLkDedup
    다중조회
End Sub

Public Sub 조회모두해제()
    Dim i As Long
    On Error Resume Next
    For i = LBound(mLkKey) To UBound(mLkKey)
        mLkKey(i) = False
        mLkVal(i) = False
    Next i
    Err.Clear
    On Error GoTo 0
    다중조회
End Sub

Private Sub 조회기준체크(ByVal 번호 As Long)
    On Error Resume Next
    mLkKey(번호) = Not mLkKey(번호)
    If mLkKey(번호) Then mLkVal(번호) = False
    Err.Clear
    On Error GoTo 0
    다중조회
End Sub

Private Sub 조회값체크(ByVal 번호 As Long)
    On Error Resume Next
    mLkVal(번호) = Not mLkVal(번호)
    If mLkVal(번호) Then mLkKey(번호) = False
    Err.Clear
    On Error GoTo 0
    다중조회
End Sub

'--- 실제 펼치기 ---
Public Sub 다중조회실행()
    Dim ws As Worksheet
    Dim 데이터 As Range
    Dim dic As Object
    Dim arr As Variant
    Dim key As String
    Dim r As Long, c As Long, i As Long, j As Long
    Dim rowCnt As Long, colCnt As Long
    Dim 최대건수 As Long
    Dim 값열() As Long, 값개수 As Long
    Dim 붙일열 As Long, 필요열 As Long
    Dim 오른쪽있음 As Boolean
    Dim 삽입함 As Boolean
    Dim 머리글 As String
    Dim oldScreen As Boolean, oldCalc As Long
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = ActiveSheet
    Set 데이터 = ws.Range(mLkAddr)
    rowCnt = 데이터.Rows.Count
    colCnt = 데이터.Columns.Count

    If rowCnt > 20000 Then
        If MsgBox(Format$(rowCnt - 1, "#,##0") & "행을 처리합니다. 시간이 걸릴 수 있습니다." & vbCrLf & _
                  "계속할까요?", vbOKCancel + vbQuestion, "다중 VLOOKUP") = vbCancel Then Exit Sub
    End If

    ' 값 열 목록
    ReDim 값열(1 To colCnt)
    For c = 1 To colCnt
        If mLkVal(c) Then
            값개수 = 값개수 + 1
            값열(값개수) = c
        End If
    Next c

    arr = 데이터.Value

    '--- 키별 건수 세어 최대 건수 파악 (중복 제거 반영) ---
    Dim 조합2 As Object
    Dim vkey2 As String

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1
    Set 조합2 = CreateObject("Scripting.Dictionary")
    조합2.CompareMode = 1

    For r = 2 To rowCnt
        key = 조회키(arr, r, colCnt)

        Dim 셀까 As Boolean
        셀까 = True
        If mLkDedup Then
            vkey2 = key & Chr$(2) & 값키(arr, r, 값열, 값개수)
            If 조합2.Exists(vkey2) Then
                셀까 = False
            Else
                조합2.Add vkey2, 1
            End If
        End If

        If 셀까 Then
            If dic.Exists(key) Then
                dic(key) = dic(key) + 1
            Else
                dic.Add key, 1
            End If
            If dic(key) > 최대건수 Then 최대건수 = dic(key)
        End If
    Next r

    If 최대건수 < 1 Then
        MsgBox "처리할 데이터가 없습니다.", vbExclamation, "다중 VLOOKUP"
        Exit Sub
    End If

    필요열 = IIf(mLkDedup, 2, 1) + 최대건수 * 값개수   ' 개수 열 + (건수 x 값열)

    If 필요열 > 200 Then
        MsgBox "필요한 열이 " & 필요열 & "개로 너무 많습니다." & vbCrLf & _
               "한 키에 최대 " & 최대건수 & "건이 있습니다." & vbCrLf & _
               "가져올 값 열을 줄여주세요.", vbExclamation, "다중 VLOOKUP"
        Exit Sub
    End If

    붙일열 = 데이터.Column + colCnt

    ' 오른쪽에 데이터가 있는지 확인
    On Error Resume Next
    오른쪽있음 = (Application.CountA(ws.Range(ws.Cells(데이터.Row, 붙일열), _
                                              ws.Cells(데이터.Row + rowCnt - 1, 붙일열 + 필요열 - 1))) > 0)
    Err.Clear
    On Error GoTo Fail

    If mLkFormula Then
        If MsgBox("수식으로 넣으면 파일이 무거워져 Excel 이 느려질 수 있습니다." & vbCrLf & _
                  "먼저 저장한 뒤 진행하시길 권합니다." & vbCrLf & vbCrLf & _
                  "계속할까요?", vbOKCancel + vbExclamation, "다중 VLOOKUP") = vbCancel Then Exit Sub
    End If

    If 오른쪽있음 Then
        If MsgBox("오른쪽에 이미 데이터가 있습니다." & vbCrLf & _
                  "빈 열 " & 필요열 & "개를 끼워 넣고 채울까요?" & vbCrLf & vbCrLf & _
                  "기존 데이터는 오른쪽으로 밀립니다." & vbCrLf & _
                  "Ctrl+Z 로 되돌릴 수 있습니다.", _
                  vbOKCancel + vbQuestion, "다중 VLOOKUP") = vbCancel Then Exit Sub
        삽입함 = True
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If 삽입함 Then
        ws.Range(ws.Columns(붙일열), ws.Columns(붙일열 + 필요열 - 1)).Insert Shift:=xlToRight
        mInsWb = ws.Parent.Name
        mInsWs = ws.Name
        mInsCol = 붙일열
        mInsCount = 필요열
        mInsOn = True
    Else
        조건열보관 ws.Range(ws.Cells(데이터.Row, 붙일열), _
                            ws.Cells(데이터.Row + rowCnt - 1, 붙일열 + 필요열 - 1))
        mInsOn = False
    End If

    '--- 머리글 ---
    j = 붙일열 - 1
    If mLkDedup Then
        j = j + 1
        With ws.Cells(데이터.Row, j)
            .Value = "총개수"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        j = j + 1
        With ws.Cells(데이터.Row, j)
            .Value = "고유개수"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    Else
        j = j + 1
        With ws.Cells(데이터.Row, j)
            .Value = "개수"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    End If

    For c = 1 To 값개수
        머리글 = Trim$(CStr(데이터.Cells(1, 값열(c)).Value))
        If Len(머리글) = 0 Then 머리글 = "값"
        For i = 1 To 최대건수
            j = j + 1
            With ws.Cells(데이터.Row, j)
                .Value = 머리글 & i
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
        Next i
    Next c

    ' 값 열의 표시 형식을 미리 입힌다
    j = 붙일열 + IIf(mLkDedup, 2, 1) - 1
    For c = 1 To 값개수
        For i = 1 To 최대건수
            j = j + 1
            ws.Cells(데이터.Row + 1, j).Resize(rowCnt - 1, 1).NumberFormat = _
                데이터.Cells(2, 값열(c)).NumberFormat
        Next i
    Next c

    '--- 본문 채우기 ---
    다중조회채우기 ws, 데이터, arr, rowCnt, colCnt, 붙일열, 최대건수, 값열, 값개수

    ws.Range(ws.Columns(붙일열), ws.Columns(붙일열 + 필요열 - 1)).AutoFit

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    If 삽입함 Then
        Application.OnUndo "다중 VLOOKUP 취소", 접두어() & "삽입열되돌리기"
    Else
        Application.OnUndo "다중 VLOOKUP 취소", 접두어() & "조건열되돌리기"
    End If
    Err.Clear
    On Error GoTo Fail

    MsgBox "펼치기를 마쳤습니다." & vbCrLf & vbCrLf & _
           "한 기준의 최대 " & 최대건수 & "개" & _
           IIf(mLkDedup, "  (같은 조합은 한 번씩만)", "  (중복 포함)") & vbCrLf & _
           "새로 만든 열 : " & 필요열 & "개" & vbCrLf & _
           "넣은 방식 : " & IIf(mLkFormula, "수식", "값") & vbCrLf & vbCrLf & _
           "Ctrl+Z 로 되돌릴 수 있습니다.", vbInformation, "다중 VLOOKUP"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
    MsgBox "펼치지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "다중 VLOOKUP"
End Sub

'--- 한 행의 기준 키 문자열 ---
Private Function 조회키(ByVal arr As Variant, ByVal r As Long, ByVal colCnt As Long) As String
    Dim c As Long
    Dim s As String
    For c = 1 To colCnt
        If mLkKey(c) Then s = s & CStr(arr(r, c)) & Chr$(1)
    Next c
    조회키 = s
End Function

'--- 본문 채우기 ---
Private Sub 다중조회채우기(ByVal ws As Worksheet, ByVal 데이터 As Range, _
                           ByVal arr As Variant, ByVal rowCnt As Long, ByVal colCnt As Long, _
                           ByVal 붙일열 As Long, ByVal 최대건수 As Long, _
                           ByRef 값열() As Long, ByVal 값개수 As Long)
    Dim 목록 As Object, 조합 As Object, 총개수 As Object
    Dim key As String, vkey As String
    Dim r As Long, c As Long, i As Long, j As Long
    Dim 행목록 As Variant
    Dim 출력() As Variant
    Dim 열자() As String
    Dim 첫행 As Long
    Dim 개수칸 As Long

    Set 목록 = CreateObject("Scripting.Dictionary")
    목록.CompareMode = 1
    Set 조합 = CreateObject("Scripting.Dictionary")
    조합.CompareMode = 1
    Set 총개수 = CreateObject("Scripting.Dictionary")
    총개수.CompareMode = 1

    개수칸 = IIf(mLkDedup, 2, 1)

    For r = 2 To rowCnt
        key = 조회키(arr, r, colCnt)

        If 총개수.Exists(key) Then
            총개수(key) = 총개수(key) + 1
        Else
            총개수.Add key, 1
        End If

        Dim 담을까 As Boolean
        담을까 = True
        If mLkDedup Then
            vkey = key & Chr$(2) & 값키(arr, r, 값열, 값개수)
            If 조합.Exists(vkey) Then
                담을까 = False
            Else
                조합.Add vkey, 1
            End If
        End If

        If 담을까 Then
            If 목록.Exists(key) Then
                목록(key) = 목록(key) & "," & r
            Else
                목록.Add key, CStr(r)
            End If
        End If
    Next r

    첫행 = 데이터.Row

    ' 값 열의 열 문자를 미리 만들어 둔다
    ReDim 열자(1 To 값개수)
    For c = 1 To 값개수
        열자(c) = 열문자(데이터.Column + 값열(c) - 1)
    Next c

    ReDim 출력(1 To rowCnt - 1, 1 To 개수칸 + 최대건수 * 값개수)

    For r = 2 To rowCnt
        key = 조회키(arr, r, colCnt)
        행목록 = Split(목록(key), ",")

        If mLkDedup Then
            출력(r - 1, 1) = 총개수(key)
            출력(r - 1, 2) = UBound(행목록) + 1
        Else
            출력(r - 1, 1) = UBound(행목록) + 1
        End If

        j = 개수칸
        For c = 1 To 값개수
            For i = 0 To 최대건수 - 1
                j = j + 1
                If i <= UBound(행목록) Then
                    If mLkFormula Then
                        출력(r - 1, j) = "=" & 열자(c) & (첫행 + CLng(행목록(i)) - 1)
                    Else
                        출력(r - 1, j) = arr(CLng(행목록(i)), 값열(c))
                    End If
                End If
            Next i
        Next c
    Next r

    ws.Range(ws.Cells(첫행 + 1, 붙일열), _
             ws.Cells(첫행 + rowCnt - 1, 붙일열 + 개수칸 + 최대건수 * 값개수 - 1)).Formula = 출력
End Sub

'--- 값 열 조합 키 (중복 제거 판정용) ---
Private Function 값키(ByVal arr As Variant, ByVal r As Long, _
                      ByRef 값열() As Long, ByVal 값개수 As Long) As String
    Dim c As Long
    Dim s As String
    For c = 1 To 값개수
        s = s & CStr(arr(r, 값열(c))) & Chr$(1)
    Next c
    값키 = s
End Function

'============================================================
'  표 끝까지 선택 (Ctrl+Shift+3)
'============================================================
Public Sub 끝까지선택()
    Dim ws As Worksheet
    Dim sel As Range
    Dim blk As Range
    Dim 끝행 As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then Exit Sub
    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "떨어져 있는 여러 범위는 한 번에 처리할 수 없습니다.", _
               vbExclamation, "표 끝까지 선택"
        Exit Sub
    End If

    On Error GoTo Fail

    Set blk = 소계행제외(sel.Cells(1, 1).CurrentRegion)
    If blk Is Nothing Then
        MsgBox "표를 찾지 못했습니다.", vbExclamation, "표 끝까지 선택"
        Exit Sub
    End If

    끝행 = blk.Row + blk.Rows.Count - 1
    If 끝행 <= sel.Row + sel.Rows.Count - 1 Then
        Application.StatusBar = "이미 표 끝까지 선택되어 있습니다. (" & 끝행 & "행)"
        상태표시줄예약
        Exit Sub
    End If

    ws.Range(ws.Cells(sel.Row, sel.Column), _
             ws.Cells(끝행, sel.Column + sel.Columns.Count - 1)).Select

    Application.StatusBar = sel.Row & "행 ~ " & 끝행 & "행 선택  (" & _
                            Format$(끝행 - sel.Row + 1, "#,##0") & "행 x " & _
                            sel.Columns.Count & "열)"
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "선택하지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "표 끝까지 선택"
End Sub

'============================================================
'  기준 열 - 새 시트 추출 (중복 행 / 고유 목록)
'============================================================
Public Sub 중복행추출()
    기준열추출 True
End Sub

Public Sub 고유목록추출()
    기준열추출 False
End Sub

Private Sub 기준열추출(ByVal 중복만 As Boolean)
    Dim ws As Worksheet, 새시트 As Worksheet
    Dim 데이터 As Range
    Dim dic As Object, 뽑음 As Object
    Dim arr As Variant, 출력() As Variant
    Dim 보임범위 As Range, 영역 As Range
    Dim 보임() As Boolean
    Dim key As String
    Dim r As Long, c As Long, i As Long, 시작 As Long
    Dim rowCnt As Long, colCnt As Long
    Dim 보이는수 As Long
    Dim 뽑은수 As Long, 출력열 As Long
    Dim 고른열 As String
    Dim 이름 As String
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = ActiveSheet
    Set 데이터 = ws.Range(mColAddr)
    rowCnt = 데이터.Rows.Count
    colCnt = 데이터.Columns.Count

    ReDim 보임(1 To rowCnt)
    On Error Resume Next
    Set 보임범위 = 데이터.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo 0

    If 보임범위 Is Nothing Then
        For r = 2 To rowCnt
            보임(r) = True
            보이는수 = 보이는수 + 1
        Next r
    Else
        For Each 영역 In 보임범위.Areas
            시작 = 영역.Row - 데이터.Row + 1
            For r = 시작 To 시작 + 영역.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not 보임(r) Then
                        보임(r) = True
                        보이는수 = 보이는수 + 1
                    End If
                End If
            Next r
        Next 영역
    End If

    If 보이는수 < 1 Then
        MsgBox "보이는 데이터가 없습니다.", vbExclamation, "기준 열 추출"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = 데이터.Value

    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = 1

    For r = 2 To rowCnt
      If 보임(r) Then
        key = 기준키(arr, r, colCnt)
        If dic.Exists(key) Then
            dic(key) = dic(key) + 1
        Else
            dic.Add key, 1
        End If
      End If
    Next r

    출력열 = 0
    For c = 1 To colCnt
        If mColPick(c) Then 출력열 = 출력열 + 1
    Next c

    If 중복만 Then
        ReDim 출력(1 To 보이는수 + 1, 1 To colCnt)
        For c = 1 To colCnt
            출력(1, c) = arr(1, c)
        Next c
        뽑은수 = 0
        For r = 2 To rowCnt
          If 보임(r) Then
            key = 기준키(arr, r, colCnt)
            If dic(key) >= 2 Then
                뽑은수 = 뽑은수 + 1
                For c = 1 To colCnt
                    출력(뽑은수 + 1, c) = arr(r, c)
                Next c
            End If
          End If
        Next r
    Else
        Set 뽑음 = CreateObject("Scripting.Dictionary")
        뽑음.CompareMode = 1
        ReDim 출력(1 To dic.Count + 1, 1 To 출력열)
        i = 0
        For c = 1 To colCnt
            If mColPick(c) Then
                i = i + 1
                출력(1, i) = arr(1, c)
            End If
        Next c
        뽑은수 = 0
        For r = 2 To rowCnt
          If 보임(r) Then
            key = 기준키(arr, r, colCnt)
            If Not 뽑음.Exists(key) Then
                뽑음.Add key, 1
                뽑은수 = 뽑은수 + 1
                i = 0
                For c = 1 To colCnt
                    If mColPick(c) Then
                        i = i + 1
                        출력(뽑은수 + 1, i) = arr(r, c)
                    End If
                Next c
            End If
          End If
        Next r
    End If

    Application.ScreenUpdating = oldScreen

    If 뽑은수 = 0 Then
        MsgBox IIf(중복만, "중복된 행이 없습니다.", "뽑을 데이터가 없습니다."), _
               vbInformation, "기준 열 추출"
        Exit Sub
    End If

    For c = 1 To colCnt
        If mColPick(c) Then
            If Len(고른열) > 0 Then 고른열 = 고른열 & " + "
            고른열 = 고른열 & Trim$(CStr(데이터.Cells(1, c).Value))
        End If
    Next c

    Application.ScreenUpdating = False
    이름 = 추출시트이름(ws.Parent, IIf(중복만, "중복_", "고유_") & ws.Name)
    Set 새시트 = ws.Parent.Worksheets.Add(Before:=ws)
    On Error Resume Next
    새시트.Name = 이름
    Err.Clear
    On Error GoTo Fail

    With 새시트.Range("A1")
        .Value = IIf(중복만, "중복 행 추출", "고유 목록") & _
                 "  |  원본 : " & ws.Name & "!" & 데이터.Address(False, False) & _
                 "  |  기준 열 : " & 고른열
        .Font.Size = 9
        .Font.Color = RGB(110, 110, 115)
    End With
    새시트.Columns("A").ColumnWidth = 12

    ' 값보다 먼저 표시 형식을 옮긴다 (00123 이 123 이 되는 것 방지)
    If 중복만 Then
        For c = 1 To colCnt
            새시트.Cells(4, c).Resize(뽑은수, 1).NumberFormat = 데이터.Cells(2, c).NumberFormat
        Next c
    Else
        i = 0
        For c = 1 To colCnt
            If mColPick(c) Then
                i = i + 1
                새시트.Cells(4, i).Resize(뽑은수, 1).NumberFormat = 데이터.Cells(2, c).NumberFormat
            End If
        Next c
    End If

    새시트.Range("A3").Resize(뽑은수 + 1, UBound(출력, 2)).Value = 출력
    With 새시트.Range("A3").Resize(1, UBound(출력, 2))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    새시트.Range("A3").Resize(뽑은수 + 1, UBound(출력, 2)).Columns.AutoFit
    새시트.Activate
    새시트.Range("A4").Select

    Application.ScreenUpdating = oldScreen

    MsgBox IIf(중복만, "중복인 행 ", "고유 목록 ") & Format$(뽑은수, "#,##0") & "행을 뽑았습니다." & vbCrLf & vbCrLf & _
           "시트 : " & 이름 & vbCrLf & _
           "기준 열 : " & 고른열 & vbCrLf & _
           "원본 : " & ws.Name & "!" & 데이터.Address(False, False) & _
           IIf(보이는수 < rowCnt - 1, "  (보이는 " & 보이는수 & "행만)", ""), _
           vbInformation, "기준 열 추출"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "뽑지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "기준 열 추출"
End Sub

Private Function 기준키(ByVal arr As Variant, ByVal r As Long, ByVal colCnt As Long) As String
    Dim c As Long
    Dim s As String
    For c = 1 To colCnt
        If mColPick(c) Then s = s & CStr(arr(r, c)) & Chr$(1)
    Next c
    기준키 = s
End Function

Private Function 추출시트이름(ByVal wb As Workbook, ByVal 기본 As String) As String
    Dim 후보 As String, 꼬리 As String
    Dim i As Long

    기본 = Replace(기본, ":", "")
    기본 = Replace(기본, "/", "")
    기본 = Replace(기본, "?", "")
    기본 = Replace(기본, "*", "")
    기본 = Replace(기본, "[", "")
    기본 = Replace(기본, "]", "")
    If Len(기본) > 31 Then 기본 = Left$(기본, 31)

    If Not 시트있음(wb, 기본) Then
        추출시트이름 = 기본
        Exit Function
    End If

    For i = 2 To 999
        꼬리 = CStr(i)
        후보 = 기본
        If Len(후보) + Len(꼬리) > 31 Then 후보 = Left$(후보, 31 - Len(꼬리))
        후보 = 후보 & 꼬리
        If Not 시트있음(wb, 후보) Then
            추출시트이름 = 후보
            Exit Function
        End If
    Next i
    추출시트이름 = "추출_" & Format$(Now, "hhnnss")
End Function

'============================================================
'  다중 VLOOKUP - 고유 목록으로 새 시트
'============================================================
Public Sub 다중조회새시트()
    Dim ws As Worksheet, 새시트 As Worksheet
    Dim 데이터 As Range
    Dim 순서 As Object, 모음 As Object, 총개수 As Object, 조합 As Object
    Dim arr As Variant, 출력() As Variant
    Dim 보임범위 As Range, 영역 As Range
    Dim 보임() As Boolean
    Dim 값열() As Long, 기준열() As Long
    Dim 값개수 As Long, 기준개수 As Long
    Dim key As String, vkey As String, 이름 As String
    Dim r As Long, c As Long, i As Long, j As Long, k As Long, 시작 As Long
    Dim rowCnt As Long, colCnt As Long
    Dim 보이는수 As Long
    Dim 최대건수 As Long, 키수 As Long
    Dim 개수칸 As Long
    Dim 담을까 As Boolean
    Dim 행목록 As Variant
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = ActiveSheet
    Set 데이터 = ws.Range(mLkAddr)
    rowCnt = 데이터.Rows.Count
    colCnt = 데이터.Columns.Count
    개수칸 = IIf(mLkDedup, 2, 1)

    ReDim 기준열(1 To colCnt)
    ReDim 값열(1 To colCnt)
    For c = 1 To colCnt
        If mLkKey(c) Then
            기준개수 = 기준개수 + 1
            기준열(기준개수) = c
        End If
        If mLkVal(c) Then
            값개수 = 값개수 + 1
            값열(값개수) = c
        End If
    Next c

    ReDim 보임(1 To rowCnt)
    On Error Resume Next
    Set 보임범위 = 데이터.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo 0
    If 보임범위 Is Nothing Then
        For r = 2 To rowCnt
            보임(r) = True
            보이는수 = 보이는수 + 1
        Next r
    Else
        For Each 영역 In 보임범위.Areas
            시작 = 영역.Row - 데이터.Row + 1
            For r = 시작 To 시작 + 영역.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not 보임(r) Then
                        보임(r) = True
                        보이는수 = 보이는수 + 1
                    End If
                End If
            Next r
        Next 영역
    End If

    If 보이는수 < 1 Then
        MsgBox "보이는 데이터가 없습니다.", vbExclamation, "다중 VLOOKUP"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error GoTo Fail
    Application.ScreenUpdating = False

    arr = 데이터.Value

    Set 순서 = CreateObject("Scripting.Dictionary")
    순서.CompareMode = 1
    Set 모음 = CreateObject("Scripting.Dictionary")
    모음.CompareMode = 1
    Set 총개수 = CreateObject("Scripting.Dictionary")
    총개수.CompareMode = 1
    Set 조합 = CreateObject("Scripting.Dictionary")
    조합.CompareMode = 1

    For r = 2 To rowCnt
      If 보임(r) Then
        key = 조회키(arr, r, colCnt)

        If 총개수.Exists(key) Then
            총개수(key) = 총개수(key) + 1
        Else
            총개수.Add key, 1
        End If

        담을까 = True
        If mLkDedup Then
            vkey = key & Chr$(2) & 값키(arr, r, 값열, 값개수)
            If 조합.Exists(vkey) Then
                담을까 = False
            Else
                조합.Add vkey, 1
            End If
        End If

        If 담을까 Then
            If 모음.Exists(key) Then
                모음(key) = 모음(key) & "," & r
            Else
                모음.Add key, CStr(r)
                키수 = 키수 + 1
                순서.Add 키수, key
            End If
            i = UBound(Split(모음(key), ",")) + 1
            If i > 최대건수 Then 최대건수 = i
        End If
      End If
    Next r

    If 기준개수 + 개수칸 + 최대건수 * 값개수 > 200 Then
        Application.ScreenUpdating = oldScreen
        MsgBox "필요한 열이 너무 많습니다." & vbCrLf & _
               "한 기준에 최대 " & 최대건수 & "개가 있습니다." & vbCrLf & _
               "가져올 값 열을 줄여주세요.", vbExclamation, "다중 VLOOKUP"
        Exit Sub
    End If

    ReDim 출력(1 To 키수 + 1, 1 To 기준개수 + 개수칸 + 최대건수 * 값개수)

    j = 0
    For i = 1 To 기준개수
        j = j + 1
        출력(1, j) = arr(1, 기준열(i))
    Next i
    If mLkDedup Then
        j = j + 1
        출력(1, j) = "총개수"
        j = j + 1
        출력(1, j) = "고유개수"
    Else
        j = j + 1
        출력(1, j) = "개수"
    End If
    For c = 1 To 값개수
        이름 = Trim$(CStr(arr(1, 값열(c))))
        If Len(이름) = 0 Then 이름 = "값"
        For i = 1 To 최대건수
            j = j + 1
            출력(1, j) = 이름 & i
        Next i
    Next c

    For k = 1 To 키수
        key = 순서(k)
        행목록 = Split(모음(key), ",")
        r = CLng(행목록(0))

        j = 0
        For i = 1 To 기준개수
            j = j + 1
            출력(k + 1, j) = arr(r, 기준열(i))
        Next i
        If mLkDedup Then
            j = j + 1
            출력(k + 1, j) = 총개수(key)
            j = j + 1
            출력(k + 1, j) = UBound(행목록) + 1
        Else
            j = j + 1
            출력(k + 1, j) = UBound(행목록) + 1
        End If

        For c = 1 To 값개수
            For i = 1 To 최대건수
                j = j + 1
                If i - 1 <= UBound(행목록) Then
                    출력(k + 1, j) = arr(CLng(행목록(i - 1)), 값열(c))
                End If
            Next i
        Next c
    Next k

    이름 = 추출시트이름(ws.Parent, "그룹_" & ws.Name)
    Set 새시트 = ws.Parent.Worksheets.Add(Before:=ws)
    On Error Resume Next
    새시트.Name = 이름
    Err.Clear
    On Error GoTo Fail

    With 새시트.Range("A1")
        .Value = "다중 VLOOKUP  |  원본 : " & ws.Name & "!" & 데이터.Address(False, False) & _
                 "  |  기준 " & 기준개수 & " / 값 " & 값개수 & _
                 IIf(mLkDedup, "  |  같은 조합은 한 번씩만", "  |  중복 포함")
        .Font.Size = 9
        .Font.Color = RGB(110, 110, 115)
    End With
    새시트.Columns("A").ColumnWidth = 12

    ' 값보다 먼저 표시 형식을 옮긴다
    j = 0
    For i = 1 To 기준개수
        j = j + 1
        새시트.Cells(4, j).Resize(키수, 1).NumberFormat = 데이터.Cells(2, 기준열(i)).NumberFormat
    Next i
    j = j + 개수칸                                  ' 개수 열은 일반 형식 그대로
    For c = 1 To 값개수
        For i = 1 To 최대건수
            j = j + 1
            새시트.Cells(4, j).Resize(키수, 1).NumberFormat = 데이터.Cells(2, 값열(c)).NumberFormat
        Next i
    Next c

    새시트.Range("A3").Resize(키수 + 1, UBound(출력, 2)).Value = 출력
    With 새시트.Range("A3").Resize(1, UBound(출력, 2))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    새시트.Range("A3").Resize(키수 + 1, UBound(출력, 2)).Columns.AutoFit
    새시트.Activate
    새시트.Range("A4").Select

    Application.ScreenUpdating = oldScreen

    MsgBox "고유 목록 " & Format$(키수, "#,##0") & "행을 만들었습니다." & vbCrLf & vbCrLf & _
           "시트 : " & 이름 & vbCrLf & _
           "한 기준의 최대 " & 최대건수 & "개" & _
           IIf(mLkDedup, "  (같은 조합은 한 번씩만)", "  (중복 포함)") & vbCrLf & _
           "원본 : " & ws.Name & "!" & 데이터.Address(False, False) & _
           IIf(보이는수 < rowCnt - 1, "  (보이는 " & 보이는수 & "행만)", ""), _
           vbInformation, "다중 VLOOKUP"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "만들지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "다중 VLOOKUP"
End Sub

'--- 조회 기준/값 토글 래퍼 ---
Public Sub 조회기준토글01()
    조회기준체크 1
End Sub
Public Sub 조회기준토글02()
    조회기준체크 2
End Sub
Public Sub 조회기준토글03()
    조회기준체크 3
End Sub
Public Sub 조회기준토글04()
    조회기준체크 4
End Sub
Public Sub 조회기준토글05()
    조회기준체크 5
End Sub
Public Sub 조회기준토글06()
    조회기준체크 6
End Sub
Public Sub 조회기준토글07()
    조회기준체크 7
End Sub
Public Sub 조회기준토글08()
    조회기준체크 8
End Sub
Public Sub 조회기준토글09()
    조회기준체크 9
End Sub
Public Sub 조회기준토글10()
    조회기준체크 10
End Sub
Public Sub 조회기준토글11()
    조회기준체크 11
End Sub
Public Sub 조회기준토글12()
    조회기준체크 12
End Sub
Public Sub 조회기준토글13()
    조회기준체크 13
End Sub
Public Sub 조회기준토글14()
    조회기준체크 14
End Sub
Public Sub 조회기준토글15()
    조회기준체크 15
End Sub
Public Sub 조회기준토글16()
    조회기준체크 16
End Sub
Public Sub 조회기준토글17()
    조회기준체크 17
End Sub
Public Sub 조회기준토글18()
    조회기준체크 18
End Sub
Public Sub 조회기준토글19()
    조회기준체크 19
End Sub
Public Sub 조회기준토글20()
    조회기준체크 20
End Sub
Public Sub 조회기준토글21()
    조회기준체크 21
End Sub
Public Sub 조회기준토글22()
    조회기준체크 22
End Sub
Public Sub 조회기준토글23()
    조회기준체크 23
End Sub
Public Sub 조회기준토글24()
    조회기준체크 24
End Sub
Public Sub 조회기준토글25()
    조회기준체크 25
End Sub
Public Sub 조회기준토글26()
    조회기준체크 26
End Sub
Public Sub 조회기준토글27()
    조회기준체크 27
End Sub
Public Sub 조회기준토글28()
    조회기준체크 28
End Sub
Public Sub 조회기준토글29()
    조회기준체크 29
End Sub
Public Sub 조회기준토글30()
    조회기준체크 30
End Sub
Public Sub 조회기준토글31()
    조회기준체크 31
End Sub
Public Sub 조회기준토글32()
    조회기준체크 32
End Sub
Public Sub 조회기준토글33()
    조회기준체크 33
End Sub
Public Sub 조회기준토글34()
    조회기준체크 34
End Sub
Public Sub 조회기준토글35()
    조회기준체크 35
End Sub
Public Sub 조회기준토글36()
    조회기준체크 36
End Sub
Public Sub 조회기준토글37()
    조회기준체크 37
End Sub
Public Sub 조회기준토글38()
    조회기준체크 38
End Sub
Public Sub 조회기준토글39()
    조회기준체크 39
End Sub
Public Sub 조회기준토글40()
    조회기준체크 40
End Sub
Public Sub 조회기준토글41()
    조회기준체크 41
End Sub
Public Sub 조회기준토글42()
    조회기준체크 42
End Sub
Public Sub 조회기준토글43()
    조회기준체크 43
End Sub
Public Sub 조회기준토글44()
    조회기준체크 44
End Sub
Public Sub 조회기준토글45()
    조회기준체크 45
End Sub
Public Sub 조회기준토글46()
    조회기준체크 46
End Sub
Public Sub 조회기준토글47()
    조회기준체크 47
End Sub
Public Sub 조회기준토글48()
    조회기준체크 48
End Sub
Public Sub 조회기준토글49()
    조회기준체크 49
End Sub
Public Sub 조회기준토글50()
    조회기준체크 50
End Sub
Public Sub 조회기준토글51()
    조회기준체크 51
End Sub
Public Sub 조회기준토글52()
    조회기준체크 52
End Sub
Public Sub 조회기준토글53()
    조회기준체크 53
End Sub
Public Sub 조회기준토글54()
    조회기준체크 54
End Sub
Public Sub 조회기준토글55()
    조회기준체크 55
End Sub
Public Sub 조회기준토글56()
    조회기준체크 56
End Sub
Public Sub 조회기준토글57()
    조회기준체크 57
End Sub
Public Sub 조회기준토글58()
    조회기준체크 58
End Sub
Public Sub 조회기준토글59()
    조회기준체크 59
End Sub
Public Sub 조회기준토글60()
    조회기준체크 60
End Sub
Public Sub 조회값토글01()
    조회값체크 1
End Sub
Public Sub 조회값토글02()
    조회값체크 2
End Sub
Public Sub 조회값토글03()
    조회값체크 3
End Sub
Public Sub 조회값토글04()
    조회값체크 4
End Sub
Public Sub 조회값토글05()
    조회값체크 5
End Sub
Public Sub 조회값토글06()
    조회값체크 6
End Sub
Public Sub 조회값토글07()
    조회값체크 7
End Sub
Public Sub 조회값토글08()
    조회값체크 8
End Sub
Public Sub 조회값토글09()
    조회값체크 9
End Sub
Public Sub 조회값토글10()
    조회값체크 10
End Sub
Public Sub 조회값토글11()
    조회값체크 11
End Sub
Public Sub 조회값토글12()
    조회값체크 12
End Sub
Public Sub 조회값토글13()
    조회값체크 13
End Sub
Public Sub 조회값토글14()
    조회값체크 14
End Sub
Public Sub 조회값토글15()
    조회값체크 15
End Sub
Public Sub 조회값토글16()
    조회값체크 16
End Sub
Public Sub 조회값토글17()
    조회값체크 17
End Sub
Public Sub 조회값토글18()
    조회값체크 18
End Sub
Public Sub 조회값토글19()
    조회값체크 19
End Sub
Public Sub 조회값토글20()
    조회값체크 20
End Sub
Public Sub 조회값토글21()
    조회값체크 21
End Sub
Public Sub 조회값토글22()
    조회값체크 22
End Sub
Public Sub 조회값토글23()
    조회값체크 23
End Sub
Public Sub 조회값토글24()
    조회값체크 24
End Sub
Public Sub 조회값토글25()
    조회값체크 25
End Sub
Public Sub 조회값토글26()
    조회값체크 26
End Sub
Public Sub 조회값토글27()
    조회값체크 27
End Sub
Public Sub 조회값토글28()
    조회값체크 28
End Sub
Public Sub 조회값토글29()
    조회값체크 29
End Sub
Public Sub 조회값토글30()
    조회값체크 30
End Sub
Public Sub 조회값토글31()
    조회값체크 31
End Sub
Public Sub 조회값토글32()
    조회값체크 32
End Sub
Public Sub 조회값토글33()
    조회값체크 33
End Sub
Public Sub 조회값토글34()
    조회값체크 34
End Sub
Public Sub 조회값토글35()
    조회값체크 35
End Sub
Public Sub 조회값토글36()
    조회값체크 36
End Sub
Public Sub 조회값토글37()
    조회값체크 37
End Sub
Public Sub 조회값토글38()
    조회값체크 38
End Sub
Public Sub 조회값토글39()
    조회값체크 39
End Sub
Public Sub 조회값토글40()
    조회값체크 40
End Sub
Public Sub 조회값토글41()
    조회값체크 41
End Sub
Public Sub 조회값토글42()
    조회값체크 42
End Sub
Public Sub 조회값토글43()
    조회값체크 43
End Sub
Public Sub 조회값토글44()
    조회값체크 44
End Sub
Public Sub 조회값토글45()
    조회값체크 45
End Sub
Public Sub 조회값토글46()
    조회값체크 46
End Sub
Public Sub 조회값토글47()
    조회값체크 47
End Sub
Public Sub 조회값토글48()
    조회값체크 48
End Sub
Public Sub 조회값토글49()
    조회값체크 49
End Sub
Public Sub 조회값토글50()
    조회값체크 50
End Sub
Public Sub 조회값토글51()
    조회값체크 51
End Sub
Public Sub 조회값토글52()
    조회값체크 52
End Sub
Public Sub 조회값토글53()
    조회값체크 53
End Sub
Public Sub 조회값토글54()
    조회값체크 54
End Sub
Public Sub 조회값토글55()
    조회값체크 55
End Sub
Public Sub 조회값토글56()
    조회값체크 56
End Sub
Public Sub 조회값토글57()
    조회값체크 57
End Sub
Public Sub 조회값토글58()
    조회값체크 58
End Sub
Public Sub 조회값토글59()
    조회값체크 59
End Sub
Public Sub 조회값토글60()
    조회값체크 60
End Sub

'============================================================
'  열 머리글을 행 데이터로  (Ctrl+Alt+B)
'    1월 2월 3월 처럼 옆으로 퍼진 열을 한 열로 길게 세운다.
'    (파워쿼리의 "열 피벗 해제" 와 같은 일을 한다)
'
'      사번    이름   1월  2월  3월
'      00123  김민준  10   20   30
'            ->
'      사번    이름   항목  값
'      00123  김민준  1월   10
'      00123  김민준  2월   20
'      00123  김민준  3월   30
'
'    팝업에서 목록 한 벌만 보여준다.
'      ● 찍은 열이 행으로 내려가고, ○ 는 그대로 남는다.
'============================================================
Public Sub 열행전환()
    Dim ws As Worksheet
    Dim 데이터 As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "열 머리글을 행 데이터로"
        Exit Sub
    End If
    Set ws = ActiveSheet

    Set 데이터 = 기준열범위(ws)
    If 데이터 Is Nothing Then Exit Sub

    If StrComp(mUpWs, ws.Name, vbTextCompare) <> 0 Or _
       StrComp(mUpAddr, 데이터.Address, vbTextCompare) <> 0 Then
        ReDim mUpVal(1 To 데이터.Columns.Count)
        mUpWs = ws.Name
        mUpAddr = 데이터.Address
        mUpBlank = True                       ' 기본 : 빈 칸도 그대로 포함
        mUpDown = False                       ' 기본 : 내릴 열 고르기
    End If

    If Len(mUpName1) = 0 Then mUpName1 = "항목"
    If Len(mUpName2) = 0 Then mUpName2 = "값"

    On Error Resume Next
    If UBound(mUpVal) <> 데이터.Columns.Count Then
        ReDim mUpVal(1 To 데이터.Columns.Count)
    End If
    Err.Clear
    On Error GoTo 0

    열행전환메뉴 데이터
End Sub

'--- 열 고르기 팝업 (목록 한 벌) ---
Private Sub 열행전환메뉴(ByVal 데이터 As Range)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim i As Long
    Dim 열개수 As Long, 풀수 As Long
    Dim 이름 As String
    Dim 원본행 As Double
    Dim 어림 As String

    On Error Resume Next
    Application.CommandBars(열행메뉴이름).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=열행메뉴이름, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    열개수 = 데이터.Columns.Count
    For i = 1 To 열개수
        If mUpVal(i) Then 풀수 = 풀수 + 1
    Next i
    원본행 = CDbl(데이터.Rows.Count - 1)

    ' 필터로 숨은 행이 있거나 빈 칸을 건너뛰면 숫자가 줄어든다
    어림 = ""
    On Error Resume Next
    If 데이터.Worksheet.FilterMode Then 어림 = "약 "
    Err.Clear
    On Error GoTo Fail
    If Not mUpBlank Then 어림 = "약 "

    메뉴제목 cb, "열 머리글을 행 데이터로        (" & 데이터.Address(False, False) & ")"
    메뉴제목 cb, "●  찍은 열이 행으로 내려갑니다.   ○  는 그대로 남습니다."

    For i = 1 To 열개수
        이름 = 열이름(데이터, i)
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mUpVal(i), "     ●  ", "     ○  ") & 이름
        btn.OnAction = 접두어() & "열행값토글" & Format$(i, "00")
        btn.Style = msoButtonCaption
        If i = 1 Then btn.BeginGroup = True
    Next i

    '--- 실행 ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  새 시트로 만들기   (" & Format$(원본행, "#,##0") & "행  ->  " & _
                  어림 & Format$(원본행 * CDbl(풀수), "#,##0") & "행)"
    btn.OnAction = 접두어() & "열행전환실행"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True
    btn.Enabled = (풀수 > 0 And 풀수 < 열개수)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     빈 칸 : " & IIf(mUpBlank, "포함  ->  건너뛰기", "건너뛰기  ->  포함")
    btn.OnAction = 접두어() & "열행빈칸전환"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     새 열 이름 : " & mUpName1 & " / " & mUpName2 & "        (클릭시 변경 가능)"
    btn.OnAction = 접두어() & "열행이름변경"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     선택 방식 : " & _
                  IIf(mUpDown, "안 내릴 열 고르기  ->  내릴 열 고르기", _
                               "내릴 열 고르기  ->  안 내릴 열 고르기")
    btn.OnAction = 접두어() & "열행방식전환"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     모두 해제        (내릴 열 " & 풀수 & "개)"
    btn.OnAction = 접두어() & "열행모두해제"
    btn.Style = msoButtonCaption

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "메뉴를 표시하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "열 머리글을 행 데이터로"
End Sub

Public Sub 열행빈칸전환()
    mUpBlank = Not mUpBlank
    열행전환
End Sub

'--- 선택 방식 바꾸기 : ● 와 ○ 를 전부 맞바꾼다 ---
Public Sub 열행방식전환()
    Dim i As Long
    On Error Resume Next
    For i = LBound(mUpVal) To UBound(mUpVal)
        mUpVal(i) = Not mUpVal(i)
    Next i
    mUpDown = Not mUpDown
    Err.Clear
    On Error GoTo 0
    열행전환
End Sub

Public Sub 열행모두해제()
    Dim i As Long
    On Error Resume Next
    For i = LBound(mUpVal) To UBound(mUpVal)
        mUpVal(i) = False
    Next i
    mUpDown = False
    Err.Clear
    On Error GoTo 0
    열행전환
End Sub

Public Sub 열행이름변경()
    Dim s As String

    s = InputBox("새로 만들 두 열의 이름을 쉼표로 적어주세요." & vbCrLf & vbCrLf & _
                 "앞쪽 = 열 머리글이 들어갈 열 이름" & vbCrLf & _
                 "뒤쪽 = 값이 들어갈 열 이름", _
                 "새 열 이름", mUpName1 & "," & mUpName2)
    If Len(Trim$(s)) = 0 Then
        열행전환
        Exit Sub
    End If

    If InStr(1, s, ",") > 0 Then
        mUpName1 = Trim$(Split(s, ",")(0))
        mUpName2 = Trim$(Split(s, ",")(1))
    Else
        mUpName1 = Trim$(s)
    End If
    If Len(mUpName1) = 0 Then mUpName1 = "항목"
    If Len(mUpName2) = 0 Then mUpName2 = "값"

    열행전환
End Sub

Private Sub 열행값체크(ByVal 번호 As Long)
    On Error Resume Next
    mUpVal(번호) = Not mUpVal(번호)
    Err.Clear
    On Error GoTo 0
    열행전환
End Sub

'------------------------------------------------------------
' 실제 변환 - 새 시트로
'------------------------------------------------------------
Public Sub 열행전환실행()
    Dim ws As Worksheet, 새시트 As Worksheet
    Dim 데이터 As Range
    Dim 보임범위 As Range, 영역 As Range
    Dim 보임() As Boolean
    Dim 고정열() As Long, 풀열() As Long
    Dim 고정개수 As Long, 풀개수 As Long
    Dim 값판() As Variant, 한열 As Variant
    Dim 머리글() As String
    Dim 버퍼() As Variant, 끝버퍼() As Variant
    Dim rowCnt As Long, 보이는수 As Long
    Dim r As Long, c As Long, i As Long, j As Long, 시작 As Long
    Dim 버퍼크기 As Long, 버퍼행 As Long, 쓴행 As Long
    Dim 결과행 As Double, 빈칸포함행 As Double
    Dim 이름 As String, 형식 As String
    Dim 빈칸수 As Double
    Dim v As Variant
    Dim 담을까 As Boolean
    Dim 열수 As Long
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldCalc As Long
    Dim 잠갔음 As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    If Len(mUpAddr) = 0 Then
        MsgBox "먼저 Ctrl+Alt+B 를 눌러 열을 골라주세요.", _
               vbInformation, "열 머리글을 행 데이터로"
        Exit Sub
    End If
    Set ws = ActiveSheet
    Set 데이터 = ws.Range(mUpAddr)
    rowCnt = 데이터.Rows.Count

    '--- ● 는 내릴 열, 나머지는 전부 기준 열 ---
    ReDim 고정열(1 To 데이터.Columns.Count)
    ReDim 풀열(1 To 데이터.Columns.Count)
    For c = 1 To 데이터.Columns.Count
        If mUpVal(c) Then
            풀개수 = 풀개수 + 1
            풀열(풀개수) = c
        Else
            고정개수 = 고정개수 + 1
            고정열(고정개수) = c
        End If
    Next c

    If 풀개수 < 1 Then
        MsgBox "행으로 내릴 열을 하나 이상 골라주세요." & vbCrLf & _
               "●  로 찍은 열이 내려갑니다.", _
               vbExclamation, "열 머리글을 행 데이터로"
        Exit Sub
    End If
    If 고정개수 < 1 Then
        MsgBox "그대로 둘 열이 하나도 없습니다." & vbCrLf & _
               "●  을 하나 이상 풀어서  ○  로 남겨주세요.", _
               vbExclamation, "열 머리글을 행 데이터로"
        Exit Sub
    End If

    '--- 보이는 행만 (필터 반영) ---
    ReDim 보임(1 To rowCnt)
    On Error Resume Next
    Set 보임범위 = 데이터.SpecialCells(xlCellTypeVisible)
    Err.Clear
    On Error GoTo Fail
    If 보임범위 Is Nothing Then
        For r = 2 To rowCnt
            보임(r) = True
            보이는수 = 보이는수 + 1
        Next r
    Else
        For Each 영역 In 보임범위.Areas
            시작 = 영역.Row - 데이터.Row + 1
            For r = 시작 To 시작 + 영역.Rows.Count - 1
                If r >= 2 And r <= rowCnt Then
                    If Not 보임(r) Then
                        보임(r) = True
                        보이는수 = 보이는수 + 1
                    End If
                End If
            Next r
        Next 영역
    End If

    If 보이는수 < 1 Then
        MsgBox "보이는 데이터가 없습니다.", vbExclamation, "열 머리글을 행 데이터로"
        Exit Sub
    End If

    '--- 필요한 열만 배열로 읽는다 (메모리 절약) ---
    Application.StatusBar = "데이터를 읽고 있습니다..."
    ReDim 값판(1 To rowCnt, 1 To 고정개수 + 풀개수)
    For i = 1 To 고정개수
        한열 = 데이터.Columns(고정열(i)).Value
        For r = 1 To rowCnt
            값판(r, i) = 한열(r, 1)
        Next r
    Next i
    ReDim 머리글(1 To 풀개수)
    For i = 1 To 풀개수
        한열 = 데이터.Columns(풀열(i)).Value
        For r = 1 To rowCnt
            값판(r, 고정개수 + i) = 한열(r, 1)
        Next r
        이름 = Trim$(CStr(데이터.Cells(1, 풀열(i)).Value))
        If Len(이름) = 0 Then 이름 = "열 " & 열문자(데이터.Column + 풀열(i) - 1)
        머리글(i) = 이름
    Next i
    한열 = Empty

    '--- 만들어질 행 수를 먼저 센다 ---
    Application.StatusBar = "만들어질 행 수를 세고 있습니다..."
    빈칸포함행 = CDbl(보이는수) * CDbl(풀개수)
    For r = 2 To rowCnt
        If 보임(r) Then
            For i = 1 To 풀개수
                v = 값판(r, 고정개수 + i)
                If IsEmpty(v) Then
                    빈칸수 = 빈칸수 + 1
                ElseIf VarType(v) = vbString Then
                    If Len(v) = 0 Then 빈칸수 = 빈칸수 + 1
                End If
            Next i
        End If
    Next r
    결과행 = 빈칸포함행 - 빈칸수
    If mUpBlank Then 결과행 = 빈칸포함행

    Application.StatusBar = False

    If 결과행 < 1 Then
        MsgBox "만들 행이 없습니다." & vbCrLf & _
               "내릴 열이 모두 비어 있습니다.", vbExclamation, "열 머리글을 행 데이터로"
        Exit Sub
    End If

    '--- 시트 한 장을 넘기면 중단 ---
    If 결과행 + 3 > CDbl(ws.Rows.Count) Then
        MsgBox "만들어질 행이 " & Format$(결과행, "#,##0") & "행으로" & vbCrLf & _
               "시트 한 장(" & Format$(ws.Rows.Count, "#,##0") & "행)을 넘습니다." & vbCrLf & vbCrLf & _
               "내릴 열을 줄이거나, 원본을 필터로 걸러서" & vbCrLf & _
               "보이는 행만 남긴 뒤 다시 실행해주세요.", _
               vbExclamation, "열 머리글을 행 데이터로"
        Exit Sub
    End If

    '--- 실행 전 확인 ---
    If MsgBox("새 시트에 " & Format$(결과행, "#,##0") & "행을 만듭니다." & vbCrLf & vbCrLf & _
              "원본 " & Format$(보이는수, "#,##0") & "행" & _
              IIf(보이는수 < rowCnt - 1, "  (보이는 행만)", "") & vbCrLf & _
              "그대로 둘 열 : " & 고정개수 & "개" & vbCrLf & _
              "행으로 내릴 열 : " & 풀개수 & "개" & vbCrLf & vbCrLf & _
              "빈 칸 포함 : " & Format$(빈칸포함행, "#,##0") & "행" & vbCrLf & _
              "빈 칸 건너뛰기 : " & Format$(빈칸포함행 - 빈칸수, "#,##0") & "행" & vbCrLf & _
              "지금 설정 : " & IIf(mUpBlank, "빈 칸 포함", "빈 칸 건너뛰기") & vbCrLf & vbCrLf & _
              IIf(결과행 > 100000, "행이 많습니다. 파일을 먼저 저장해두시는 편이 안전합니다." & vbCrLf & vbCrLf, "") & _
              "진행하시겠습니까?", _
              vbOKCancel + vbQuestion, "열 머리글을 행 데이터로") <> vbOK Then
        Exit Sub
    End If

    '--- 잠그기 ---
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldCalc = Application.Calculation
    잠갔음 = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    열수 = 고정개수 + 2

    '--- 새 시트 ---
    이름 = 추출시트이름(ws.Parent, "행변환_" & ws.Name)
    Set 새시트 = ws.Parent.Worksheets.Add(Before:=ws)
    On Error Resume Next
    새시트.Name = 이름
    Err.Clear
    On Error GoTo Fail

    With 새시트.Range("A1")
        .Value = "열 머리글을 행 데이터로  |  원본 : " & ws.Name & "!" & 데이터.Address(False, False) & _
                 "  |  내린 열 " & 풀개수 & "개" & _
                 IIf(mUpBlank, "  |  빈 칸 포함", "  |  빈 칸 건너뛰기")
        .Font.Size = 9
        .Font.Color = RGB(110, 110, 115)
    End With
    새시트.Columns("A").ColumnWidth = 12

    '--- 값보다 먼저 표시 형식을 옮긴다 ---
    For i = 1 To 고정개수
        새시트.Cells(4, i).Resize(CLng(결과행), 1).NumberFormat = _
            데이터.Cells(2, 고정열(i)).NumberFormat
    Next i
    형식 = 데이터.Cells(2, 풀열(1)).NumberFormat
    For i = 2 To 풀개수
        If 데이터.Cells(2, 풀열(i)).NumberFormat <> 형식 Then
            형식 = "General"
            Exit For
        End If
    Next i
    새시트.Cells(4, 고정개수 + 1).Resize(CLng(결과행), 1).NumberFormat = "@"
    새시트.Cells(4, 열수).Resize(CLng(결과행), 1).NumberFormat = 형식

    '--- 머리글 ---
    ReDim 버퍼(1 To 1, 1 To 열수)
    For i = 1 To 고정개수
        버퍼(1, i) = 값판(1, i)
    Next i
    버퍼(1, 고정개수 + 1) = mUpName1
    버퍼(1, 열수) = mUpName2
    새시트.Range("A3").Resize(1, 열수).Value = 버퍼
    With 새시트.Range("A3").Resize(1, 열수)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    '--- 본문을 덩어리로 나눠 쓴다 ---
    버퍼크기 = 50000
    If 버퍼크기 > CLng(결과행) Then 버퍼크기 = CLng(결과행)
    ReDim 버퍼(1 To 버퍼크기, 1 To 열수)
    버퍼행 = 0
    쓴행 = 0

    For r = 2 To rowCnt
      If 보임(r) Then
        For i = 1 To 풀개수
            v = 값판(r, 고정개수 + i)
            담을까 = True
            If Not mUpBlank Then
                If IsEmpty(v) Then
                    담을까 = False
                ElseIf VarType(v) = vbString Then
                    If Len(v) = 0 Then 담을까 = False
                End If
            End If

            If 담을까 Then
                버퍼행 = 버퍼행 + 1
                For c = 1 To 고정개수
                    버퍼(버퍼행, c) = 값판(r, c)
                Next c
                버퍼(버퍼행, 고정개수 + 1) = 머리글(i)
                버퍼(버퍼행, 열수) = v

                If 버퍼행 = 버퍼크기 Then
                    새시트.Cells(4 + 쓴행, 1).Resize(버퍼크기, 열수).Value = 버퍼
                    쓴행 = 쓴행 + 버퍼크기
                    버퍼행 = 0
                    Application.StatusBar = "만들고 있습니다...  " & _
                        Format$(쓴행, "#,##0") & " / " & Format$(결과행, "#,##0") & "행"
                End If
            End If
        Next i
      End If
    Next r

    '--- 남은 것 ---
    If 버퍼행 > 0 Then
        ReDim 끝버퍼(1 To 버퍼행, 1 To 열수)
        For j = 1 To 버퍼행
            For c = 1 To 열수
                끝버퍼(j, c) = 버퍼(j, c)
            Next c
        Next j
        새시트.Cells(4 + 쓴행, 1).Resize(버퍼행, 열수).Value = 끝버퍼
        쓴행 = 쓴행 + 버퍼행
    End If

    새시트.Range("A3").Resize(1, 열수).Columns.AutoFit
    새시트.Activate
    새시트.Range("A4").Select

    '--- 풀기 ---
    Application.Calculation = oldCalc
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    잠갔음 = False
    Application.StatusBar = False

    MsgBox Format$(쓴행, "#,##0") & "행을 만들었습니다." & vbCrLf & vbCrLf & _
           "시트 : " & 이름 & vbCrLf & _
           "열 : 고정 " & 고정개수 & "개 + " & mUpName1 & " + " & mUpName2 & vbCrLf & _
           "원본 : " & ws.Name & "!" & 데이터.Address(False, False) & _
           IIf(보이는수 < rowCnt - 1, "  (보이는 " & Format$(보이는수, "#,##0") & "행만)", ""), _
           vbInformation, "열 머리글을 행 데이터로"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    If 잠갔음 Then
        Application.Calculation = oldCalc
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        Application.ScreenUpdating = oldScreen
    End If
    Application.StatusBar = False
    MsgBox "만들지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc & vbCrLf & vbCrLf & _
           "행이 너무 많으면 내릴 열을 줄이거나" & vbCrLf & _
           "원본을 나눠서 실행해주세요.", vbExclamation, "열 머리글을 행 데이터로"
End Sub

'--- 내릴 열 토글 래퍼 ---
Public Sub 열행값토글01()
    열행값체크 1
End Sub
Public Sub 열행값토글02()
    열행값체크 2
End Sub
Public Sub 열행값토글03()
    열행값체크 3
End Sub
Public Sub 열행값토글04()
    열행값체크 4
End Sub
Public Sub 열행값토글05()
    열행값체크 5
End Sub
Public Sub 열행값토글06()
    열행값체크 6
End Sub
Public Sub 열행값토글07()
    열행값체크 7
End Sub
Public Sub 열행값토글08()
    열행값체크 8
End Sub
Public Sub 열행값토글09()
    열행값체크 9
End Sub
Public Sub 열행값토글10()
    열행값체크 10
End Sub
Public Sub 열행값토글11()
    열행값체크 11
End Sub
Public Sub 열행값토글12()
    열행값체크 12
End Sub
Public Sub 열행값토글13()
    열행값체크 13
End Sub
Public Sub 열행값토글14()
    열행값체크 14
End Sub
Public Sub 열행값토글15()
    열행값체크 15
End Sub
Public Sub 열행값토글16()
    열행값체크 16
End Sub
Public Sub 열행값토글17()
    열행값체크 17
End Sub
Public Sub 열행값토글18()
    열행값체크 18
End Sub
Public Sub 열행값토글19()
    열행값체크 19
End Sub
Public Sub 열행값토글20()
    열행값체크 20
End Sub
Public Sub 열행값토글21()
    열행값체크 21
End Sub
Public Sub 열행값토글22()
    열행값체크 22
End Sub
Public Sub 열행값토글23()
    열행값체크 23
End Sub
Public Sub 열행값토글24()
    열행값체크 24
End Sub
Public Sub 열행값토글25()
    열행값체크 25
End Sub
Public Sub 열행값토글26()
    열행값체크 26
End Sub
Public Sub 열행값토글27()
    열행값체크 27
End Sub
Public Sub 열행값토글28()
    열행값체크 28
End Sub
Public Sub 열행값토글29()
    열행값체크 29
End Sub
Public Sub 열행값토글30()
    열행값체크 30
End Sub
Public Sub 열행값토글31()
    열행값체크 31
End Sub
Public Sub 열행값토글32()
    열행값체크 32
End Sub
Public Sub 열행값토글33()
    열행값체크 33
End Sub
Public Sub 열행값토글34()
    열행값체크 34
End Sub
Public Sub 열행값토글35()
    열행값체크 35
End Sub
Public Sub 열행값토글36()
    열행값체크 36
End Sub
Public Sub 열행값토글37()
    열행값체크 37
End Sub
Public Sub 열행값토글38()
    열행값체크 38
End Sub
Public Sub 열행값토글39()
    열행값체크 39
End Sub
Public Sub 열행값토글40()
    열행값체크 40
End Sub
Public Sub 열행값토글41()
    열행값체크 41
End Sub
Public Sub 열행값토글42()
    열행값체크 42
End Sub
Public Sub 열행값토글43()
    열행값체크 43
End Sub
Public Sub 열행값토글44()
    열행값체크 44
End Sub
Public Sub 열행값토글45()
    열행값체크 45
End Sub
Public Sub 열행값토글46()
    열행값체크 46
End Sub
Public Sub 열행값토글47()
    열행값체크 47
End Sub
Public Sub 열행값토글48()
    열행값체크 48
End Sub
Public Sub 열행값토글49()
    열행값체크 49
End Sub
Public Sub 열행값토글50()
    열행값체크 50
End Sub
Public Sub 열행값토글51()
    열행값체크 51
End Sub
Public Sub 열행값토글52()
    열행값체크 52
End Sub
Public Sub 열행값토글53()
    열행값체크 53
End Sub
Public Sub 열행값토글54()
    열행값체크 54
End Sub
Public Sub 열행값토글55()
    열행값체크 55
End Sub
Public Sub 열행값토글56()
    열행값체크 56
End Sub
Public Sub 열행값토글57()
    열행값체크 57
End Sub
Public Sub 열행값토글58()
    열행값체크 58
End Sub
Public Sub 열행값토글59()
    열행값체크 59
End Sub
Public Sub 열행값토글60()
    열행값체크 60
End Sub

'--- 시작 값이 어떤 종류인지 ---
Private Function 채우기종류(ByVal s As String) As String
    If IsNumeric(s) Then
        채우기종류 = "숫자"
    ElseIf IsDate(s) Then
        채우기종류 = "날짜"
    ElseIf 요일번호(s, False) > 0 Then
        채우기종류 = "요일"
    ElseIf 요일번호(s, True) > 0 Then
        채우기종류 = "요일긴"
    ElseIf 꼬리숫자(s, "월") > 0 Then
        채우기종류 = "달"
    ElseIf 꼬리숫자(s, "분기") > 0 Then
        채우기종류 = "분기"
    ElseIf 꼬리숫자(s, "주") > 0 Then
        채우기종류 = "주"
    ElseIf 꼬리숫자(s, "일") > 0 Then
        채우기종류 = "일"
    ElseIf 앞숫자(s) > 0 Then
        채우기종류 = "앞숫자"
    Else
        채우기종류 = "복사"
    End If
End Function

'--- n 번째 값 ---
Private Function 채우기값(ByVal 시작값 As String, ByVal 종류 As String, _
                          ByVal n As Long, ByVal 증가 As Double) As Variant
    Dim i As Long
    Dim 꼬리 As String

    Select Case 종류
        Case "숫자"
            채우기값 = CDbl(시작값) + n * 증가

        Case "날짜"
            채우기값 = CDate(시작값) + n * 증가

        Case "요일"
            i = ((요일번호(시작값, False) - 1 + n) Mod 7) + 1
            채우기값 = Mid$("일월화수목금토", i, 1)

        Case "요일긴"
            i = ((요일번호(시작값, True) - 1 + n) Mod 7) + 1
            채우기값 = Mid$("일월화수목금토", i, 1) & "요일"

        Case "달"
            i = 꼬리숫자(시작값, "월") + n
            i = ((i - 1) Mod 12) + 1
            채우기값 = i & "월"

        Case "분기"
            i = 꼬리숫자(시작값, "분기") + n
            i = ((i - 1) Mod 4) + 1
            채우기값 = i & "분기"

        Case "주"
            채우기값 = (꼬리숫자(시작값, "주") + n) & "주"

        Case "일"
            채우기값 = (꼬리숫자(시작값, "일") + n) & "일"

        Case "앞숫자"
            꼬리 = Mid$(시작값, Len(CStr(앞숫자(시작값))) + 1)
            채우기값 = (앞숫자(시작값) + n) & 꼬리

        Case Else
            채우기값 = 시작값
    End Select
End Function

'--- 요일 문자열 -> 1(일) ~ 7(토), 아니면 0 ---
Private Function 요일번호(ByVal s As String, ByVal 긴것 As Boolean) As Long
    Dim 표 As String
    Dim i As Long
    s = Trim$(s)
    표 = "일월화수목금토"
    If 긴것 Then
        If Len(s) <> 3 Then Exit Function
        If Right$(s, 2) <> "요일" Then Exit Function
        s = Left$(s, 1)
    Else
        If Len(s) <> 1 Then Exit Function
    End If
    i = InStr(표, s)
    If i > 0 Then 요일번호 = i
End Function

'--- "3월" -> 3,  "2분기" -> 2.  아니면 0 ---
Private Function 꼬리숫자(ByVal s As String, ByVal 꼬리 As String) As Long
    Dim 앞 As String
    s = Trim$(s)
    If Len(s) <= Len(꼬리) Then Exit Function
    If Right$(s, Len(꼬리)) <> 꼬리 Then Exit Function
    앞 = Left$(s, Len(s) - Len(꼬리))
    If Not IsNumeric(앞) Then Exit Function
    If CDbl(앞) <> Int(CDbl(앞)) Then Exit Function
    꼬리숫자 = CLng(앞)
End Function

'--- "1차", "2팀" 처럼 앞이 숫자면 그 숫자. 아니면 0 ---
Private Function 앞숫자(ByVal s As String) As Long
    Dim i As Long
    Dim 숫자 As String
    s = Trim$(s)
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then
            숫자 = 숫자 & Mid$(s, i, 1)
        Else
            Exit For
        End If
    Next i
    If Len(숫자) = 0 Then Exit Function
    If Len(숫자) = Len(s) Then Exit Function
    앞숫자 = CLng(숫자)
End Function

'============================================================
'  가운데 정렬 / 기본 (Ctrl+Alt+Z)
'============================================================
Public Sub 가운데정렬()
    Dim rng As Range
    Dim 지금 As Variant
    Dim 가운데로 As Boolean
    Dim errNo As Long, errDesc As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "셀을 선택한 뒤 실행해주세요.", vbExclamation, "가운데 정렬"
        Exit Sub
    End If
    Set rng = Selection

    On Error GoTo Fail

    지금 = rng.HorizontalAlignment
    가운데로 = True
    If Not IsNull(지금) Then
        If 지금 = xlCenter Then 가운데로 = False
    End If

    정렬보관 rng

    If 가운데로 Then
        rng.HorizontalAlignment = xlCenter
    Else
        rng.HorizontalAlignment = xlGeneral
    End If

    If mAlignOn Then
        On Error Resume Next
        Application.OnUndo "정렬 취소", 접두어() & "정렬되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = IIf(가운데로, "가운데 정렬했습니다.", "기본 정렬로 되돌렸습니다.") & _
                            "  (" & rng.Address(False, False) & ")" & _
                            IIf(mAlignOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "정렬을 바꾸지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "가운데 정렬"
End Sub

'============================================================
'  선택 영역의 가운데로 (Ctrl+Alt+S)
'    셀을 합치지 않고 걸친 것처럼 보이게 한다.
'    병합과 달리 정렬 / 필터 / 피벗이 정상 동작한다.
'============================================================
Public Sub 가운데걸치기()
    Dim rng As Range
    Dim 지금 As Variant
    Dim 걸치기로 As Boolean
    Dim errNo As Long, errDesc As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "셀을 선택한 뒤 실행해주세요.", vbExclamation, "선택 영역의 가운데로"
        Exit Sub
    End If
    Set rng = Selection

    If rng.Columns.Count < 2 Then
        MsgBox "가로로 2칸 이상 선택해주세요." & vbCrLf & vbCrLf & _
               "맨 왼쪽 칸의 글자가 선택 영역 가운데에 걸쳐 보입니다." & vbCrLf & _
               "셀을 합치지 않아 정렬·필터·피벗이 정상 동작합니다.", _
               vbExclamation, "선택 영역의 가운데로"
        Exit Sub
    End If

    On Error GoTo Fail

    지금 = rng.HorizontalAlignment
    걸치기로 = True
    If Not IsNull(지금) Then
        If 지금 = xlCenterAcrossSelection Then 걸치기로 = False
    End If

    정렬보관 rng

    If 걸치기로 Then
        rng.HorizontalAlignment = xlCenterAcrossSelection
    Else
        rng.HorizontalAlignment = xlGeneral
    End If

    If mAlignOn Then
        On Error Resume Next
        Application.OnUndo "정렬 취소", 접두어() & "정렬되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = IIf(걸치기로, "선택 영역의 가운데로 걸쳤습니다. (셀은 합치지 않음)", _
                                          "기본 정렬로 되돌렸습니다.") & _
                            IIf(mAlignOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "바꾸지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "선택 영역의 가운데로"
End Sub

'--- 정렬 되돌리기용 보관 ---
Private Sub 정렬보관(ByVal rg As Range)
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

Public Sub 정렬되돌리기()
    Dim rng As Range
    Dim r As Long, c As Long
    Dim oldScreen As Boolean

    If Not mAlignOn Or Len(mAlignAddr) = 0 Then
        MsgBox "되돌릴 정렬이 없습니다.", vbInformation, "정렬"
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

    Application.StatusBar = "정렬을 되돌렸습니다."
    상태표시줄예약
End Sub

'============================================================
'  위 칸 값으로 빈 칸 채우기 (Ctrl+Alt+A)
'    병합을 푼 표나 시스템에서 내려온 데이터의 빈 칸을
'    바로 위 값으로 흘려 넣는다.
'============================================================
Public Sub 위칸채우기()
    Dim ws As Worksheet
    Dim rng As Range
    Dim blk As Range
    Dim arr As Variant
    Dim r As Long, c As Long
    Dim 채운수 As Long
    Dim oldScreen As Boolean, oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "빈 칸을 채울 범위를 선택한 뒤 실행해주세요.", vbExclamation, "위 칸 값 채우기"
        Exit Sub
    End If

    If Selection.Cells.CountLarge >= 2 And Selection.Areas.Count = 1 Then
        Set rng = Selection
    Else
        Set blk = 소계행제외(ActiveCell.CurrentRegion)
        If blk Is Nothing Then
            MsgBox "표를 찾지 못했습니다.", vbExclamation, "위 칸 값 채우기"
            Exit Sub
        End If
        Set rng = ws.Range(ws.Cells(blk.Row + 1, ActiveCell.Column), _
                           ws.Cells(blk.Row + blk.Rows.Count - 1, ActiveCell.Column))
    End If

    If rng Is Nothing Then
        MsgBox "채울 데이터가 없습니다.", vbExclamation, "위 칸 값 채우기"
        Exit Sub
    End If
    If rng.Rows.Count < 2 Then
        MsgBox "2행 이상을 선택해주세요.", vbExclamation, "위 칸 값 채우기"
        Exit Sub
    End If
    If rng.Cells.CountLarge > 200000 Then
        MsgBox "범위가 너무 큽니다. 좁혀서 다시 실행해주세요.", vbExclamation, "위 칸 값 채우기"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    조건열보관 rng
    arr = rng.Value

    For c = 1 To rng.Columns.Count
        For r = 2 To rng.Rows.Count
            If Len(CStr(arr(r, c))) = 0 Then
                If Len(CStr(arr(r - 1, c))) > 0 Then
                    arr(r, c) = arr(r - 1, c)
                    채운수 = 채운수 + 1
                End If
            End If
        Next r
    Next c

    If 채운수 > 0 Then rng.Value = arr

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    If 채운수 = 0 Then
        MsgBox "채울 빈 칸이 없습니다." & vbCrLf & _
               "범위 : " & rng.Address(False, False), vbInformation, "위 칸 값 채우기"
        Exit Sub
    End If

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "빈 칸 채우기 취소", 접두어() & "조건열되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "빈 칸 " & Format$(채운수, "#,##0") & "개를 위 칸 값으로 채웠습니다. (" & _
                            rng.Address(False, False) & ")" & _
                            IIf(mCondOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
    MsgBox "채우지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "위 칸 값 채우기"
End Sub

'============================================================
'  머리글 위 행 제거 (Ctrl+Alt+E)
'    텍스트로 내보내기 전에 머리글 위쪽 잡동사니를 지운다.
'============================================================
Public Sub 머리글위제거()
    Dim ws As Worksheet
    Dim blk As Range
    Dim 머리글행 As Long
    Dim 지울수 As Long
    Dim 미리보기 As String
    Dim i As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    On Error GoTo Fail

    Set blk = 소계행제외(ActiveCell.CurrentRegion)
    If blk Is Nothing Then
        MsgBox "표를 찾지 못했습니다.", vbExclamation, "머리글 위 행 제거"
        Exit Sub
    End If

    머리글행 = blk.Row
    지울수 = 머리글행 - 1

    If 지울수 < 1 Then
        MsgBox "머리글이 이미 1행에 있습니다." & vbCrLf & _
               "지울 행이 없습니다.", vbInformation, "머리글 위 행 제거"
        Exit Sub
    End If

    For i = 1 To 지울수
        If i > 5 Then
            미리보기 = 미리보기 & "   ..." & vbCrLf
            Exit For
        End If
        미리보기 = 미리보기 & "   " & i & "행 : " & 행미리보기(ws, i) & vbCrLf
    Next i

    If MsgBox("머리글 위 " & 지울수 & "행을 지웁니다." & vbCrLf & vbCrLf & _
              미리보기 & vbCrLf & _
              "머리글로 본 행 : " & 머리글행 & "행  (" & 행미리보기(ws, 머리글행) & ")" & vbCrLf & vbCrLf & _
              "Ctrl+Z 로 되돌릴 수 있습니다." & vbCrLf & vbCrLf & _
              "계속할까요?", vbOKCancel + vbQuestion, "머리글 위 행 제거") = vbCancel Then Exit Sub

    머리글위보관 ws, 지울수

    ws.Range(ws.Rows(1), ws.Rows(지울수)).Delete Shift:=xlUp

    If mDelOn Then
        On Error Resume Next
        Application.OnUndo "머리글 위 행 제거 취소", 접두어() & "머리글위되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = 지울수 & "행을 지웠습니다. 머리글이 1행으로 올라왔습니다." & _
                            IIf(mDelOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    MsgBox "지우지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "머리글 위 행 제거"
End Sub

Private Function 행미리보기(ByVal ws As Worksheet, ByVal 행 As Long) As String
    Dim c As Long
    Dim s As String
    Dim v As String
    On Error GoTo Done
    For c = 1 To 12
        v = Trim$(CStr(ws.Cells(행, c).Value))
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
    If Len(s) = 0 Then s = "(빈 행)"
    행미리보기 = s
End Function

'============================================================
'  두 열 대조 (Ctrl+Alt+D)
'    A열 각 값이 B열에 있는지, B열 각 값이 A열에 있는지를
'    1/0 과 개수로 나란히 표시한다.
'============================================================
Public Sub 두열대조()
    Dim ws As Worksheet
    Dim 데이터 As Range

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "두 열 대조"
        Exit Sub
    End If
    Set ws = ActiveSheet

    '--- Ctrl 로 두 곳을 잡았으면 무엇을 할지 고르게 한다 ---
    If TypeName(Selection) = "Range" Then
        If Selection.Areas.Count = 2 Then
            mCmpDirA = Selection.Areas(1).Address
            mCmpDirB = Selection.Areas(2).Address
            mCmpDirWs = ws.Name
            두곳메뉴
            Exit Sub
        End If
    End If

    Set 데이터 = 기준열범위(ws)
    If 데이터 Is Nothing Then Exit Sub

    If StrComp(mCmpWs, ws.Name, vbTextCompare) <> 0 Or _
       StrComp(mCmpAddr, 데이터.Address, vbTextCompare) <> 0 Then
        mCmpWs = ws.Name
        mCmpAddr = 데이터.Address
        mCmpA = 0
        mCmpB = 0
        mCmpFormula = True
    End If

    두열대조메뉴 데이터
End Sub

'--- Ctrl 로 잡은 두 범위를 바로 대조 ---
Private Sub 대조바로실행(ByVal rngA As Range, ByVal rngB As Range)
    Dim A이름 As String, B이름 As String
    Dim a As Range, b As Range

    Set a = rngA
    Set b = rngB
    비교범위정리 a, b
    Set rngA = a
    Set rngB = b

    If rngA.Columns.Count > 1 Or rngB.Columns.Count > 1 Then
        MsgBox "한 열씩 잡아 주세요." & vbCrLf & _
               "Ctrl 을 누른 채 비교할 두 열(또는 두 구간)을 하나씩 선택합니다.", _
               vbExclamation, "두 열 대조"
        Exit Sub
    End If

    A이름 = 범위머리글(rngA)
    B이름 = 범위머리글(rngB)

    대조핵심 rngA, rngB, A이름, B이름
End Sub

'--- 범위 위쪽 칸에서 머리글을 찾는다 ---
Private Function 범위머리글(ByVal rg As Range) As String
    Dim s As String
    On Error GoTo Done
    If rg.Row > 1 Then s = Trim$(CStr(rg.Cells(1, 1).Offset(-1, 0).Value))
Done:
    If Len(s) = 0 Then s = Split(rg.Cells(1, 1).Address(True, False), "$")(0) & "열"
    범위머리글 = s
End Function

Private Sub 두열대조메뉴(ByVal 데이터 As Range)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim i As Long

    On Error Resume Next
    Application.CommandBars(두열대조메뉴이름).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=두열대조메뉴이름, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    메뉴제목 cb, "두 열 비교  -  아래에서 방식을 고르세요"

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "①  대조 대상 열  (하나)"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    For i = 1 To 데이터.Columns.Count
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mCmpA = i, "     ■  ", "     □  ") & 열이름(데이터, i)
        btn.OnAction = 접두어() & "대조A토글" & Format$(i, "00")
        btn.Style = msoButtonCaption
    Next i

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "②  찾을 목록 열  (하나)"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    For i = 1 To 데이터.Columns.Count
        Set btn = cb.Controls.Add(Type:=msoControlButton)
        btn.Caption = IIf(mCmpB = i, "     ●  ", "     ○  ") & 열이름(데이터, i)
        btn.OnAction = 접두어() & "대조B토글" & Format$(i, "00")
        btn.Style = msoButtonCaption
    Next i

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     결과 : " & IIf(mCmpFormula, "수식으로 (COUNTIF, 자동 갱신)", "값으로 (빠름)")
    btn.OnAction = 접두어() & "대조방식전환"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  목록에 있는지 대조   (결과 4열, 양방향)"
    btn.OnAction = 접두어() & "두열대조실행"
    btn.Style = msoButtonCaption
    btn.Enabled = (mCmpA > 0 And mCmpB > 0 And mCmpA <> mCmpB)

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  같은 행끼리 비교   (다른 행만 표시 + 색칠)"
    btn.OnAction = 접두어() & "행비교실행"
    btn.Style = msoButtonCaption
    btn.Enabled = (mCmpA > 0 And mCmpB > 0 And mCmpA <> mCmpB)

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "메뉴를 표시하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "두 열 대조"
End Sub

Public Sub 대조방식전환()
    mCmpFormula = Not mCmpFormula
    두열대조
End Sub

Private Sub 대조A체크(ByVal 번호 As Long)
    mCmpA = 번호
    If mCmpB = 번호 Then mCmpB = 0
    두열대조
End Sub

Private Sub 대조B체크(ByVal 번호 As Long)
    mCmpB = 번호
    If mCmpA = 번호 Then mCmpA = 0
    두열대조
End Sub

Public Sub 두열대조실행()
    Dim ws As Worksheet
    Dim 데이터 As Range
    Dim rngA As Range, rngB As Range

    On Error GoTo Fail
    Set ws = ActiveSheet
    Set 데이터 = ws.Range(mCmpAddr)

    Set rngA = 데이터.Columns(mCmpA).Offset(1, 0).Resize(데이터.Rows.Count - 1, 1)
    Set rngB = 데이터.Columns(mCmpB).Offset(1, 0).Resize(데이터.Rows.Count - 1, 1)

    대조핵심 rngA, rngB, 열이름(데이터, mCmpA), 열이름(데이터, mCmpB)
    Exit Sub
Fail:
    MsgBox "대조하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "두 열 대조"
End Sub

'--- 실제 대조 : 두 범위를 받아 결과 4열을 오른쪽에 넣는다 ---
Private Sub 대조핵심(ByVal rngA As Range, ByVal rngB As Range, _
                     ByVal A이름 As String, ByVal B이름 As String)
    Dim ws As Worksheet
    Dim dicA As Object, dicB As Object
    Dim arrA As Variant, arrB As Variant
    Dim 출력() As Variant
    Dim key As String
    Dim r As Long
    Dim nA As Long, nB As Long, 최대 As Long
    Dim A개수 As Long, B개수 As Long
    Dim A없음 As Long, B없음 As Long, 양쪽 As Long
    Dim 붙일열 As Long, 첫행 As Long
    Dim 오른쪽 As Long
    Dim 삽입함 As Boolean, 있음 As Boolean
    Dim A열자 As String, B열자 As String
    Dim A절대 As String, B절대 As String
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail

    Set ws = rngA.Worksheet
    nA = rngA.Rows.Count
    nB = rngB.Rows.Count
    최대 = nA
    If nB > 최대 Then 최대 = nB

    If 최대 > 100000 Then
        If MsgBox(Format$(최대, "#,##0") & "행을 대조합니다. 시간이 걸릴 수 있습니다." & vbCrLf & _
                  "계속할까요?", vbOKCancel + vbQuestion, "두 열 대조") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    arrA = 한열배열(rngA)
    arrB = 한열배열(rngB)

    Set dicA = CreateObject("Scripting.Dictionary")
    dicA.CompareMode = 1
    Set dicB = CreateObject("Scripting.Dictionary")
    dicB.CompareMode = 1

    For r = 1 To nA
        key = Trim$(CStr(arrA(r)))
        If Len(key) > 0 Then
            A개수 = A개수 + 1
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
            B개수 = B개수 + 1
            If dicB.Exists(key) Then
                dicB(key) = dicB(key) + 1
            Else
                dicB.Add key, 1
            End If
        End If
    Next r

    If A개수 = 0 Or B개수 = 0 Then
        Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
        MsgBox "대조할 값이 없습니다." & vbCrLf & _
               A이름 & " : " & A개수 & "행 / " & B이름 & " : " & B개수 & "행", _
               vbExclamation, "두 열 대조"
        Exit Sub
    End If

    ' 결과 위치 : 두 범위 중 오른쪽 끝에서 한 칸 띄고
    오른쪽 = rngA.Column
    If rngB.Column > 오른쪽 Then 오른쪽 = rngB.Column
    붙일열 = 오른쪽 + 2
    첫행 = rngA.Row
    If rngB.Row < 첫행 Then 첫행 = rngB.Row

    ' 주소 문자열을 미리 만들어 둔다 (행마다 .Address 를 부르면 매우 느림)
    A열자 = 열문자(rngA.Column)
    B열자 = 열문자(rngB.Column)
    A절대 = rngA.Address(True, True)
    B절대 = rngB.Address(True, True)

    ReDim 출력(1 To 최대 + 1, 1 To 4)
    출력(1, 1) = B이름 & "에존재"
    출력(1, 2) = B이름 & "개수"
    출력(1, 3) = A이름 & "에존재"
    출력(1, 4) = A이름 & "개수"

    For r = 1 To 최대
        If r <= nA Then
            key = Trim$(CStr(arrA(r)))
            If Len(key) > 0 Then
                If mCmpFormula Then
                    출력(r + 1, 1) = "=IF(COUNTIF(" & B절대 & "," & A열자 & (rngA.Row + r - 1) & ")>0,1,0)"
                    출력(r + 1, 2) = "=COUNTIF(" & B절대 & "," & A열자 & (rngA.Row + r - 1) & ")"
                Else
                    출력(r + 1, 1) = IIf(dicB.Exists(key), 1, 0)
                    출력(r + 1, 2) = IIf(dicB.Exists(key), dicB(key), 0)
                End If
                If dicB.Exists(key) Then
                    양쪽 = 양쪽 + 1
                Else
                    A없음 = A없음 + 1
                End If
            End If
        End If

        If r <= nB Then
            key = Trim$(CStr(arrB(r)))
            If Len(key) > 0 Then
                If mCmpFormula Then
                    출력(r + 1, 3) = "=IF(COUNTIF(" & A절대 & "," & B열자 & (rngB.Row + r - 1) & ")>0,1,0)"
                    출력(r + 1, 4) = "=COUNTIF(" & A절대 & "," & B열자 & (rngB.Row + r - 1) & ")"
                Else
                    출력(r + 1, 3) = IIf(dicA.Exists(key), 1, 0)
                    출력(r + 1, 4) = IIf(dicA.Exists(key), dicA(key), 0)
                End If
                If Not dicA.Exists(key) Then B없음 = B없음 + 1
            End If
        End If
    Next r

    ' 오른쪽에 값이 있으면 물어보고 열 4개를 끼워 넣는다
    삽입함 = False
    On Error Resume Next
    있음 = (Application.CountA(ws.Range(ws.Cells(첫행 - 1, 붙일열), _
                                        ws.Cells(첫행 + 최대 - 1, 붙일열 + 3))) > 0)
    Err.Clear
    On Error GoTo Fail

    If 있음 Then
        Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
        If MsgBox("결과를 넣을 자리에 이미 데이터가 있습니다." & vbCrLf & _
                  "빈 열 4개를 끼워 넣고 채울까요?" & vbCrLf & vbCrLf & _
                  "기존 데이터는 오른쪽으로 밀립니다." & vbCrLf & _
                  "Ctrl+Z 로 되돌릴 수 있습니다.", _
                  vbOKCancel + vbQuestion, "두 열 대조") = vbCancel Then Exit Sub
        Application.ScreenUpdating = False
        ws.Range(ws.Columns(붙일열), ws.Columns(붙일열 + 3)).Insert Shift:=xlToRight
        삽입함 = True
        mInsWb = ws.Parent.Name
        mInsWs = ws.Name
        mInsCol = 붙일열
        mInsCount = 4
        mInsOn = True
    Else
        조건열보관 ws.Range(ws.Cells(첫행 - 1, 붙일열), ws.Cells(첫행 + 최대 - 1, 붙일열 + 3))
        mInsOn = False
    End If

    ws.Range(ws.Cells(첫행 - 1, 붙일열), _
             ws.Cells(첫행 + 최대 - 1, 붙일열 + 3)).Formula = 출력

    With ws.Range(ws.Cells(첫행 - 1, 붙일열), ws.Cells(첫행 - 1, 붙일열 + 3))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Columns(붙일열), ws.Columns(붙일열 + 3)).AutoFit

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    On Error Resume Next
    If 삽입함 Then
        Application.OnUndo "두 열 대조 취소", 접두어() & "삽입열되돌리기"
    ElseIf mCondOn Then
        Application.OnUndo "두 열 대조 취소", 접두어() & "조건열되돌리기"
    End If
    Err.Clear
    On Error GoTo Fail

    MsgBox "대조를 마쳤습니다." & vbCrLf & vbCrLf & _
           "양쪽에 다 있는 행 : " & Format$(양쪽, "#,##0") & "행" & vbCrLf & vbCrLf & _
           A이름 & " " & Format$(A개수, "#,##0") & "행 중 " & B이름 & "에 없는 것 : " & _
           Format$(A없음, "#,##0") & "행" & vbCrLf & _
           B이름 & " " & Format$(B개수, "#,##0") & "행 중 " & A이름 & "에 없는 것 : " & _
           Format$(B없음, "#,##0") & "행" & vbCrLf & vbCrLf & _
           "넣은 방식 : " & IIf(mCmpFormula, "수식 (COUNTIF)", "값") & vbCrLf & _
           "0 만 필터하면 없는 것만 볼 수 있습니다." & vbCrLf & _
           "Ctrl+Z 로 되돌릴 수 있습니다.", vbInformation, "두 열 대조"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "대조하지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "두 열 대조"
End Sub

'--- 한 열 범위를 1차원 배열로 ---
Private Function 한열배열(ByVal rg As Range) As Variant
    Dim v As Variant
    Dim a() As Variant
    Dim r As Long

    If rg.Rows.Count = 1 Then
        ReDim a(1 To 1)
        a(1) = rg.Cells(1, 1).Value
        한열배열 = a
        Exit Function
    End If

    v = rg.Value
    ReDim a(1 To rg.Rows.Count)
    For r = 1 To rg.Rows.Count
        a(r) = v(r, 1)
    Next r
    한열배열 = a
End Function

'============================================================
'  단축키 하나 테스트
'    특정 단축키가 안 먹을 때 원인을 가른다.
'    임시로 그 키에 알림만 걸어 보고, 확인 후 원래대로 돌린다.
'============================================================


'--- 두 열 대조 토글 래퍼 ---
Public Sub 대조A토글01()
    대조A체크 1
End Sub
Public Sub 대조A토글02()
    대조A체크 2
End Sub
Public Sub 대조A토글03()
    대조A체크 3
End Sub
Public Sub 대조A토글04()
    대조A체크 4
End Sub
Public Sub 대조A토글05()
    대조A체크 5
End Sub
Public Sub 대조A토글06()
    대조A체크 6
End Sub
Public Sub 대조A토글07()
    대조A체크 7
End Sub
Public Sub 대조A토글08()
    대조A체크 8
End Sub
Public Sub 대조A토글09()
    대조A체크 9
End Sub
Public Sub 대조A토글10()
    대조A체크 10
End Sub
Public Sub 대조A토글11()
    대조A체크 11
End Sub
Public Sub 대조A토글12()
    대조A체크 12
End Sub
Public Sub 대조A토글13()
    대조A체크 13
End Sub
Public Sub 대조A토글14()
    대조A체크 14
End Sub
Public Sub 대조A토글15()
    대조A체크 15
End Sub
Public Sub 대조A토글16()
    대조A체크 16
End Sub
Public Sub 대조A토글17()
    대조A체크 17
End Sub
Public Sub 대조A토글18()
    대조A체크 18
End Sub
Public Sub 대조A토글19()
    대조A체크 19
End Sub
Public Sub 대조A토글20()
    대조A체크 20
End Sub
Public Sub 대조A토글21()
    대조A체크 21
End Sub
Public Sub 대조A토글22()
    대조A체크 22
End Sub
Public Sub 대조A토글23()
    대조A체크 23
End Sub
Public Sub 대조A토글24()
    대조A체크 24
End Sub
Public Sub 대조A토글25()
    대조A체크 25
End Sub
Public Sub 대조A토글26()
    대조A체크 26
End Sub
Public Sub 대조A토글27()
    대조A체크 27
End Sub
Public Sub 대조A토글28()
    대조A체크 28
End Sub
Public Sub 대조A토글29()
    대조A체크 29
End Sub
Public Sub 대조A토글30()
    대조A체크 30
End Sub
Public Sub 대조A토글31()
    대조A체크 31
End Sub
Public Sub 대조A토글32()
    대조A체크 32
End Sub
Public Sub 대조A토글33()
    대조A체크 33
End Sub
Public Sub 대조A토글34()
    대조A체크 34
End Sub
Public Sub 대조A토글35()
    대조A체크 35
End Sub
Public Sub 대조A토글36()
    대조A체크 36
End Sub
Public Sub 대조A토글37()
    대조A체크 37
End Sub
Public Sub 대조A토글38()
    대조A체크 38
End Sub
Public Sub 대조A토글39()
    대조A체크 39
End Sub
Public Sub 대조A토글40()
    대조A체크 40
End Sub
Public Sub 대조A토글41()
    대조A체크 41
End Sub
Public Sub 대조A토글42()
    대조A체크 42
End Sub
Public Sub 대조A토글43()
    대조A체크 43
End Sub
Public Sub 대조A토글44()
    대조A체크 44
End Sub
Public Sub 대조A토글45()
    대조A체크 45
End Sub
Public Sub 대조A토글46()
    대조A체크 46
End Sub
Public Sub 대조A토글47()
    대조A체크 47
End Sub
Public Sub 대조A토글48()
    대조A체크 48
End Sub
Public Sub 대조A토글49()
    대조A체크 49
End Sub
Public Sub 대조A토글50()
    대조A체크 50
End Sub
Public Sub 대조A토글51()
    대조A체크 51
End Sub
Public Sub 대조A토글52()
    대조A체크 52
End Sub
Public Sub 대조A토글53()
    대조A체크 53
End Sub
Public Sub 대조A토글54()
    대조A체크 54
End Sub
Public Sub 대조A토글55()
    대조A체크 55
End Sub
Public Sub 대조A토글56()
    대조A체크 56
End Sub
Public Sub 대조A토글57()
    대조A체크 57
End Sub
Public Sub 대조A토글58()
    대조A체크 58
End Sub
Public Sub 대조A토글59()
    대조A체크 59
End Sub
Public Sub 대조A토글60()
    대조A체크 60
End Sub
Public Sub 대조B토글01()
    대조B체크 1
End Sub
Public Sub 대조B토글02()
    대조B체크 2
End Sub
Public Sub 대조B토글03()
    대조B체크 3
End Sub
Public Sub 대조B토글04()
    대조B체크 4
End Sub
Public Sub 대조B토글05()
    대조B체크 5
End Sub
Public Sub 대조B토글06()
    대조B체크 6
End Sub
Public Sub 대조B토글07()
    대조B체크 7
End Sub
Public Sub 대조B토글08()
    대조B체크 8
End Sub
Public Sub 대조B토글09()
    대조B체크 9
End Sub
Public Sub 대조B토글10()
    대조B체크 10
End Sub
Public Sub 대조B토글11()
    대조B체크 11
End Sub
Public Sub 대조B토글12()
    대조B체크 12
End Sub
Public Sub 대조B토글13()
    대조B체크 13
End Sub
Public Sub 대조B토글14()
    대조B체크 14
End Sub
Public Sub 대조B토글15()
    대조B체크 15
End Sub
Public Sub 대조B토글16()
    대조B체크 16
End Sub
Public Sub 대조B토글17()
    대조B체크 17
End Sub
Public Sub 대조B토글18()
    대조B체크 18
End Sub
Public Sub 대조B토글19()
    대조B체크 19
End Sub
Public Sub 대조B토글20()
    대조B체크 20
End Sub
Public Sub 대조B토글21()
    대조B체크 21
End Sub
Public Sub 대조B토글22()
    대조B체크 22
End Sub
Public Sub 대조B토글23()
    대조B체크 23
End Sub
Public Sub 대조B토글24()
    대조B체크 24
End Sub
Public Sub 대조B토글25()
    대조B체크 25
End Sub
Public Sub 대조B토글26()
    대조B체크 26
End Sub
Public Sub 대조B토글27()
    대조B체크 27
End Sub
Public Sub 대조B토글28()
    대조B체크 28
End Sub
Public Sub 대조B토글29()
    대조B체크 29
End Sub
Public Sub 대조B토글30()
    대조B체크 30
End Sub
Public Sub 대조B토글31()
    대조B체크 31
End Sub
Public Sub 대조B토글32()
    대조B체크 32
End Sub
Public Sub 대조B토글33()
    대조B체크 33
End Sub
Public Sub 대조B토글34()
    대조B체크 34
End Sub
Public Sub 대조B토글35()
    대조B체크 35
End Sub
Public Sub 대조B토글36()
    대조B체크 36
End Sub
Public Sub 대조B토글37()
    대조B체크 37
End Sub
Public Sub 대조B토글38()
    대조B체크 38
End Sub
Public Sub 대조B토글39()
    대조B체크 39
End Sub
Public Sub 대조B토글40()
    대조B체크 40
End Sub
Public Sub 대조B토글41()
    대조B체크 41
End Sub
Public Sub 대조B토글42()
    대조B체크 42
End Sub
Public Sub 대조B토글43()
    대조B체크 43
End Sub
Public Sub 대조B토글44()
    대조B체크 44
End Sub
Public Sub 대조B토글45()
    대조B체크 45
End Sub
Public Sub 대조B토글46()
    대조B체크 46
End Sub
Public Sub 대조B토글47()
    대조B체크 47
End Sub
Public Sub 대조B토글48()
    대조B체크 48
End Sub
Public Sub 대조B토글49()
    대조B체크 49
End Sub
Public Sub 대조B토글50()
    대조B체크 50
End Sub
Public Sub 대조B토글51()
    대조B체크 51
End Sub
Public Sub 대조B토글52()
    대조B체크 52
End Sub
Public Sub 대조B토글53()
    대조B체크 53
End Sub
Public Sub 대조B토글54()
    대조B체크 54
End Sub
Public Sub 대조B토글55()
    대조B체크 55
End Sub
Public Sub 대조B토글56()
    대조B체크 56
End Sub
Public Sub 대조B토글57()
    대조B체크 57
End Sub
Public Sub 대조B토글58()
    대조B체크 58
End Sub
Public Sub 대조B토글59()
    대조B체크 59
End Sub
Public Sub 대조B토글60()
    대조B체크 60
End Sub

'============================================================
'  자동 채우기 (Ctrl+Alt+A)
'    1,2,3 / 월,화,수 / 날짜 같은 패턴을 인식해 표 끝까지 늘린다.
'    씨앗이 2칸 이상이어야 패턴을 읽는다. (Excel 자동 채우기와 같음)
'============================================================
Public Sub 자동채우기()
    Dim ws As Worksheet
    Dim sel As Range
    Dim blk As Range
    Dim 대상 As Range
    Dim c As Range
    Dim 끝행 As Long
    Dim 덮어쓸 As Long
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "자동 채우기"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "씨앗이 될 칸을 선택한 뒤 실행해주세요.", vbExclamation, "자동 채우기"
        Exit Sub
    End If

    Set sel = Selection
    If sel.Areas.Count > 1 Then
        MsgBox "떨어져 있는 여러 범위는 한 번에 처리할 수 없습니다.", _
               vbExclamation, "자동 채우기"
        Exit Sub
    End If

    On Error GoTo Fail

    Set blk = 소계행제외(sel.Cells(1, 1).CurrentRegion)
    If blk Is Nothing Then
        MsgBox "표를 찾지 못했습니다.", vbExclamation, "자동 채우기"
        Exit Sub
    End If

    끝행 = blk.Row + blk.Rows.Count - 1
    If 끝행 <= sel.Row + sel.Rows.Count - 1 Then
        MsgBox "아래로 채울 행이 없습니다." & vbCrLf & _
               "표 마지막 행에 있거나 표를 찾지 못했습니다.", _
               vbExclamation, "자동 채우기"
        Exit Sub
    End If

    Set 대상 = ws.Range(ws.Cells(sel.Row, sel.Column), _
                        ws.Cells(끝행, sel.Column + sel.Columns.Count - 1))

    '--- 씨앗 아래에 이미 값이 있으면 확인 ---
    덮어쓸 = 0
    On Error Resume Next
    For Each c In sel.Columns
        덮어쓸 = 덮어쓸 + Application.CountA(ws.Range( _
                    ws.Cells(sel.Row + sel.Rows.Count, c.Column), ws.Cells(끝행, c.Column)))
    Next c
    Err.Clear
    On Error GoTo Fail

    If 덮어쓸 > 0 Then
        If MsgBox("아래쪽에 이미 값이 " & Format$(덮어쓸, "#,##0") & "칸 있습니다." & vbCrLf & _
                  "덮어쓸까요?" & vbCrLf & vbCrLf & _
                  "채울 범위 : " & (sel.Row + sel.Rows.Count) & "행 ~ " & 끝행 & "행" & vbCrLf & _
                  "Ctrl+Z 로 되돌릴 수 있습니다.", _
                  vbOKCancel + vbExclamation, "자동 채우기") = vbCancel Then Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    조건열보관 대상

    sel.AutoFill Destination:=대상, Type:=xlFillDefault

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "자동 채우기 취소", 접두어() & "조건열되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "패턴을 이어 " & (sel.Row + sel.Rows.Count) & "행 ~ " & 끝행 & _
                            "행까지 채웠습니다." & _
                            IIf(sel.Rows.Count < 2, "   (씨앗이 1칸이라 같은 값이 복사됩니다)", "") & _
                            IIf(mCondOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "채우지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc & vbCrLf & vbCrLf & _
           "패턴을 읽으려면 씨앗이 2칸 이상이어야 합니다.", _
           vbExclamation, "자동 채우기"
End Sub

'--- 머리글 위 행 제거 : 되돌리기용 보관 ---
Private Sub 머리글위보관(ByVal ws As Worksheet, ByVal 행수 As Long)
    Dim r As Long, c As Long, nC As Long
    Dim vF() As Variant, vN() As Variant, vC() As Variant
    Dim vI() As Variant, vB() As Variant, vA() As Variant, vH() As Variant

    mDelOn = False
    On Error GoTo Fail

    nC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If nC < 1 Then Exit Sub
    If nC > 200 Then nC = 200
    If 행수 * nC > 20000 Then Exit Sub

    ReDim vF(1 To 행수, 1 To nC)
    ReDim vN(1 To 행수, 1 To nC)
    ReDim vC(1 To 행수, 1 To nC)
    ReDim vI(1 To 행수, 1 To nC)
    ReDim vB(1 To 행수, 1 To nC)
    ReDim vA(1 To 행수, 1 To nC)
    ReDim vH(1 To 행수)

    For r = 1 To 행수
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
    mDelRows = 행수
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
' 머리글 위 행 제거 되돌리기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 머리글위되돌리기()
    Dim ws As Worksheet
    Dim r As Long, c As Long
    Dim oldScreen As Boolean

    If Not mDelOn Or mDelRows < 1 Then
        MsgBox "되돌릴 행 제거가 없습니다.", vbInformation, "머리글 위 행 제거"
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

    Application.StatusBar = "지웠던 행을 되돌렸습니다."
    상태표시줄예약
End Sub

'============================================================
'  빠른 채우기 (Ctrl+Alt+A)  -  Ctrl+E 업그레이드
'    커서 위쪽을 예시로 보고, 커서부터 표 끝까지 알아서 채운다.
'    Ctrl+E 와 달리 채울 범위를 직접 잡을 필요가 없다.
'============================================================
Public Sub 빠른채우기()
    Dim ws As Worksheet
    Dim 커서 As Range
    Dim blk As Range
    Dim 대상 As Range
    Dim 예시행 As Long, 시작행 As Long, 끝행 As Long
    Dim 열 As Long
    Dim 채운수 As Long, r As Long
    Dim 덮어쓸 As Long
    Dim 안내 As String
    Dim 채운범위 As Range
    Dim 결과값 As Variant
    Dim oldScreen As Boolean
    Dim oldCalc As Long
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "빠른 채우기"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "채우기를 시작할 칸을 선택한 뒤 실행해주세요.", vbExclamation, "빠른 채우기"
        Exit Sub
    End If

    Set 커서 = ActiveCell
    열 = 커서.Column
    시작행 = 커서.Row

    On Error GoTo Fail

    Set blk = 소계행제외(커서.CurrentRegion)
    If blk Is Nothing Then
        MsgBox "표를 찾지 못했습니다." & vbCrLf & _
               "옆 열에 데이터가 있어야 규칙을 읽을 수 있습니다.", _
               vbExclamation, "빠른 채우기"
        Exit Sub
    End If

    끝행 = blk.Row + blk.Rows.Count - 1
    If 끝행 < 시작행 Then
        MsgBox "채울 행이 없습니다.", vbExclamation, "빠른 채우기"
        Exit Sub
    End If

    '--- 커서 위쪽에 예시가 몇 칸 있는지 ---
    예시행 = 0
    For r = 시작행 - 1 To blk.Row Step -1
        If Len(CStr(ws.Cells(r, 열).Value)) = 0 Then Exit For
        예시행 = 예시행 + 1
    Next r

    If 예시행 = 0 Then
        MsgBox "커서 바로 위에 예시가 없습니다." & vbCrLf & vbCrLf & _
               "원하는 결과를 1~2칸 직접 입력한 뒤" & vbCrLf & _
               "그 아래 칸에 커서를 두고 다시 실행해주세요.", _
               vbExclamation, "빠른 채우기"
        Exit Sub
    End If

    '--- 채울 자리에 이미 값이 있는지 ---
    덮어쓸 = 0
    On Error Resume Next
    덮어쓸 = Application.CountA(ws.Range(ws.Cells(시작행, 열), ws.Cells(끝행, 열)))
    Err.Clear
    On Error GoTo Fail

    If 예시행 = 1 Or 덮어쓸 > 0 Then
        안내 = ""
        If 예시행 = 1 Then
            안내 = 안내 & "· 예시가 1칸입니다. 2칸 이상이면 규칙을 더 정확히 읽습니다." & vbCrLf
        End If
        If 덮어쓸 > 0 Then
            안내 = 안내 & "· 채울 자리에 이미 값이 " & Format$(덮어쓸, "#,##0") & "칸 있습니다. 덮어씁니다." & vbCrLf
        End If
        If MsgBox(안내 & vbCrLf & _
                  "채울 범위 : " & 시작행 & "행 ~ " & 끝행 & "행" & vbCrLf & _
                  "Ctrl+Z 로 되돌릴 수 있습니다." & vbCrLf & vbCrLf & _
                  "계속할까요?", vbOKCancel + vbQuestion, "빠른 채우기") = vbCancel Then Exit Sub
    End If

    Set 대상 = ws.Range(ws.Cells(시작행 - 예시행, 열), ws.Cells(끝행, 열))

    oldScreen = Application.ScreenUpdating
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    Err.Clear
    On Error GoTo Fail

    조건열보관 ws.Range(ws.Cells(시작행, 열), ws.Cells(끝행, 열))

    On Error Resume Next
    대상.FlashFill
    Err.Clear
    On Error GoTo Fail

    ' 실제로 채워졌는지 세어 본다
    For r = 시작행 To 끝행
        If Len(CStr(ws.Cells(r, 열).Value)) > 0 Then 채운수 = 채운수 + 1
    Next r

    ' FlashFill 은 Excel 내장 동작이라 되돌리기 스택을 자기 것으로 덮어쓴다.
    ' 채워진 값을 VBA 로 다시 써서 스택 주인을 되돌린다.
    If 채운수 > 0 Then
        Set 채운범위 = ws.Range(ws.Cells(시작행, 열), ws.Cells(끝행, 열))
        결과값 = 채운범위.Value
        채운범위.Value = 결과값
    End If

    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen

    If 채운수 = 0 Then
        MsgBox "규칙을 읽지 못했습니다." & vbCrLf & vbCrLf & _
               "· 예시를 2칸 이상 넣어보세요." & vbCrLf & _
               "· 옆 열에 원본 데이터가 있어야 합니다." & vbCrLf & _
               "· 규칙이 애매하면 빠른 채우기가 동작하지 않습니다.", _
               vbExclamation, "빠른 채우기"
        Exit Sub
    End If

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "빠른 채우기 취소", 접두어() & "조건열되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "예시 " & 예시행 & "칸을 보고 " & 시작행 & "행 ~ " & 끝행 & _
                            "행을 채웠습니다.  (" & Format$(채운수, "#,##0") & "칸)" & _
                            IIf(mCondOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreen
    MsgBox "채우지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "빠른 채우기"
End Sub

'============================================================
'  서식 복사 / 붙여넣기 (Ctrl+Alt+X)
'    한 번 누르면 그 칸의 서식을 기억하고,
'    다른 곳을 선택해 다시 누르면 그 서식을 입힌다.
'    붓 아이콘을 마우스로 누를 필요가 없다.
'============================================================
Public Sub 서식복사()
    Dim ws As Worksheet
    Dim rng As Range
    Dim 원본 As Range
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "워크시트에서 실행해주세요.", vbExclamation, "서식 복사"
        Exit Sub
    End If
    Set ws = ActiveSheet

    If TypeName(Selection) <> "Range" Then
        MsgBox "셀을 선택한 뒤 실행해주세요.", vbExclamation, "서식 복사"
        Exit Sub
    End If
    Set rng = Selection

    On Error GoTo Fail

    '--- 기억해 둔 서식이 없으면 : 복사 ---
    If Not mFmtSrcOn Then
        mFmtSrcWb = ws.Parent.Name
        mFmtSrcWs = ws.Name
        mFmtSrcAddr = rng.Address
        mFmtSrcOn = True

        ' Esc 로 취소할 수 있게 하고, 40초 뒤 자동 해제
        On Error Resume Next
        Application.OnKey "{ESC}", 접두어() & "서식복사취소"
        mFmtSrcTime = Now + TimeSerial(0, 0, 서식복사_유지초)
        Application.OnTime mFmtSrcTime, 접두어() & "서식복사시간종료"
        Err.Clear
        On Error GoTo Fail

        ' 이 안내는 4초 뒤 지우지 않고 붙여넣을 때까지 남긴다
        Application.StatusBar = "[ 서식 복사 대기 중 ]  " & rng.Address(False, False) & _
                                "   ->  붙여넣을 곳을 선택하고 Ctrl+Alt+X" & _
                                "      ( Esc 취소 · " & 서식복사_유지초 & "초 뒤 자동 해제 )"
        Exit Sub
    End If

    '--- 기억해 둔 서식이 있으면 : 붙여넣기 ---
    Set 원본 = Nothing
    On Error Resume Next
    Set 원본 = Application.Workbooks(mFmtSrcWb).Worksheets(mFmtSrcWs).Range(mFmtSrcAddr)
    Err.Clear
    On Error GoTo Fail

    If 원본 Is Nothing Then
        mFmtSrcOn = False
        MsgBox "기억해 둔 서식을 찾지 못했습니다." & vbCrLf & _
               "다시 복사할 곳을 선택하고 실행해주세요.", vbExclamation, "서식 복사"
        Exit Sub
    End If

    ' 같은 자리를 다시 누르면 취소
    If StrComp(mFmtSrcWs, ws.Name, vbTextCompare) = 0 And _
       StrComp(mFmtSrcAddr, rng.Address, vbTextCompare) = 0 Then
        서식복사해제
        Application.StatusBar = "서식 복사를 취소했습니다."
        상태표시줄예약
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    조건열보관 rng

    원본.Copy
    rng.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    서식복사해제
    Application.ScreenUpdating = oldScreen

    If mCondOn Then
        On Error Resume Next
        Application.OnUndo "서식 붙여넣기 취소", 접두어() & "조건열되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    Application.StatusBar = "서식을 붙여넣었습니다. (" & rng.Address(False, False) & ")" & _
                            IIf(mCondOn, "  |  Ctrl+Z 로 되돌리기", "")
    상태표시줄예약
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
    서식복사해제
    MsgBox "서식을 처리하지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "서식 복사"
End Sub

'============================================================
'  같은 행끼리 비교 (Ctrl+Alt+F)
'    두 열을 행 단위로 맞대어 값이 다른 행만 표시한다.
'    예) 부서 vs 발령부서  ->  부서를 옮기는 사람만 걸러낸다.
'============================================================
Private Sub 행비교바로(ByVal rngA As Range, ByVal rngB As Range)
    Dim a As Range, b As Range

    Set a = rngA
    Set b = rngB
    비교범위정리 a, b
    Set rngA = a
    Set rngB = b

    If rngA.Columns.Count > 1 Or rngB.Columns.Count > 1 Then
        MsgBox "한 열씩 잡아 주세요." & vbCrLf & _
               "Ctrl 을 누른 채 비교할 두 열을 하나씩 선택합니다.", _
               vbExclamation, "같은 행끼리 비교"
        Exit Sub
    End If
    행비교핵심 rngA, rngB, 범위머리글(rngA), 범위머리글(rngB)
End Sub

Public Sub 행비교실행()
    Dim ws As Worksheet
    Dim 데이터 As Range
    Dim rngA As Range, rngB As Range

    On Error GoTo Fail
    Set ws = ActiveSheet
    Set 데이터 = ws.Range(mCmpAddr)

    Set rngA = 데이터.Columns(mCmpA).Offset(1, 0).Resize(데이터.Rows.Count - 1, 1)
    Set rngB = 데이터.Columns(mCmpB).Offset(1, 0).Resize(데이터.Rows.Count - 1, 1)

    행비교핵심 rngA, rngB, 열이름(데이터, mCmpA), 열이름(데이터, mCmpB)
    Exit Sub
Fail:
    MsgBox "비교하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "같은 행끼리 비교"
End Sub

'--- 실제 비교 ---
Private Sub 행비교핵심(ByVal rngA As Range, ByVal rngB As Range, _
                       ByVal A이름 As String, ByVal B이름 As String)
    Dim ws As Worksheet
    Dim arrA As Variant, arrB As Variant
    Dim 결과() As Variant
    Dim r As Long, n As Long
    Dim vA As String, vB As String
    Dim 다른수 As Long, 빈칸 As Long
    Dim 붙일열 As Long, 첫행 As Long
    Dim 결과범위 As Range, 색범위 As Range
    Dim 색칠 As Range
    Dim 삽입함 As Boolean, 있음 As Boolean
    Dim oldScreen As Boolean
    Dim errNo As Long, errDesc As String

    On Error GoTo Fail
    Set ws = rngA.Worksheet

    n = rngA.Rows.Count
    If rngB.Rows.Count < n Then n = rngB.Rows.Count
    If rngA.Rows.Count <> rngB.Rows.Count Then
        If MsgBox("두 범위의 행 수가 다릅니다." & vbCrLf & vbCrLf & _
                  A이름 & " : " & rngA.Rows.Count & "행" & vbCrLf & _
                  B이름 & " : " & rngB.Rows.Count & "행" & vbCrLf & vbCrLf & _
                  "짧은 쪽인 " & n & "행까지만 비교합니다. 계속할까요?", _
                  vbOKCancel + vbExclamation, "같은 행끼리 비교") = vbCancel Then Exit Sub
    End If

    If n < 1 Then
        MsgBox "비교할 행이 없습니다.", vbExclamation, "같은 행끼리 비교"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    arrA = 한열배열(rngA)
    arrB = 한열배열(rngB)

    ReDim 결과(1 To n + 1, 1 To 1)
    결과(1, 1) = A이름 & " ≠ " & B이름

    For r = 1 To n
        vA = Trim$(CStr(arrA(r)))
        vB = Trim$(CStr(arrB(r)))

        If Len(vA) = 0 And Len(vB) = 0 Then
            결과(r + 1, 1) = 0
        ElseIf StrComp(vA, vB, vbTextCompare) = 0 Then
            결과(r + 1, 1) = 0
        Else
            결과(r + 1, 1) = 1
            다른수 = 다른수 + 1
            If Len(vA) = 0 Or Len(vB) = 0 Then 빈칸 = 빈칸 + 1
            If 색칠 Is Nothing Then
                Set 색칠 = Application.Union(rngA.Cells(r, 1), rngB.Cells(r, 1))
            Else
                Set 색칠 = Application.Union(색칠, rngA.Cells(r, 1), rngB.Cells(r, 1))
            End If
        End If
    Next r

    ' 결과 열 위치 : 두 범위 중 오른쪽 끝에서 한 칸 띄고
    붙일열 = rngA.Column
    If rngB.Column > 붙일열 Then 붙일열 = rngB.Column
    붙일열 = 붙일열 + 2
    첫행 = rngA.Row
    If rngB.Row < 첫행 Then 첫행 = rngB.Row

    ' 오른쪽에 값이 있으면 물어보고 열을 끼워 넣는다
    삽입함 = False
    On Error Resume Next
    있음 = (Application.CountA(ws.Range(ws.Cells(첫행 - 1, 붙일열), _
                                        ws.Cells(첫행 + n - 1, 붙일열))) > 0)
    Err.Clear
    On Error GoTo Fail

    If 있음 Then
        Application.ScreenUpdating = oldScreen
        If MsgBox("결과를 넣을 자리에 이미 데이터가 있습니다." & vbCrLf & _
                  "빈 열 하나를 끼워 넣고 채울까요?" & vbCrLf & vbCrLf & _
                  "기존 데이터는 오른쪽으로 밀립니다." & vbCrLf & _
                  "Ctrl+Z 로 되돌릴 수 있습니다.", _
                  vbOKCancel + vbQuestion, "같은 행끼리 비교") = vbCancel Then Exit Sub
        Application.ScreenUpdating = False
        ws.Columns(붙일열).Insert Shift:=xlToRight
        삽입함 = True
    End If

    Set 결과범위 = ws.Range(ws.Cells(첫행 - 1, 붙일열), ws.Cells(첫행 + n - 1, 붙일열))
    Set 색범위 = Application.Union(rngA.Resize(n, 1), rngB.Resize(n, 1))

    행비교보관 결과범위, 색범위, 삽입함, 붙일열

    결과범위.Value = 결과
    With ws.Cells(첫행 - 1, 붙일열)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    결과범위.HorizontalAlignment = xlCenter
    ws.Columns(붙일열).AutoFit

    If Not 색칠 Is Nothing Then 색칠.Interior.Color = RGB(255, 199, 206)

    Application.ScreenUpdating = oldScreen

    If mDifOn Then
        On Error Resume Next
        Application.OnUndo "같은 행끼리 비교 취소", 접두어() & "행비교되돌리기"
        Err.Clear
        On Error GoTo Fail
    End If

    MsgBox "비교를 마쳤습니다." & vbCrLf & vbCrLf & _
           A이름 & " vs " & B이름 & vbCrLf & vbCrLf & _
           "비교한 행 : " & Format$(n, "#,##0") & "행" & vbCrLf & _
           "값이 다른 행 : " & Format$(다른수, "#,##0") & "행" & _
           IIf(빈칸 > 0, "   (한쪽이 빈 칸 " & 빈칸 & "행 포함)", "") & vbCrLf & vbCrLf & _
           "다른 행은 분홍색으로 칠했습니다." & vbCrLf & _
           "결과 열에서 1 만 필터하면 모아 볼 수 있습니다." & vbCrLf & _
           "Ctrl+Z 로 되돌릴 수 있습니다.", vbInformation, "같은 행끼리 비교"
    Exit Sub

Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = oldScreen
    MsgBox "비교하지 못했습니다." & vbCrLf & _
           "오류 " & errNo & ": " & errDesc, vbExclamation, "같은 행끼리 비교"
End Sub

'--- 되돌리기용 보관 : 결과 열 값 + 비교 범위 색 ---
Private Sub 행비교보관(ByVal 결과범위 As Range, ByVal 색범위 As Range, _
                       ByVal 삽입함 As Boolean, ByVal 붙일열 As Long)
    Dim ar As Range, cel As Range
    Dim v() As Variant, vC() As Variant, vI() As Variant
    Dim i As Long, n As Long

    mDifOn = False
    On Error GoTo Fail

    If 결과범위.Rows.Count > 60000 Then Exit Sub

    ReDim v(1 To 결과범위.Rows.Count)
    For i = 1 To 결과범위.Rows.Count
        v(i) = 결과범위.Cells(i, 1).Formula
    Next i

    n = 0
    For Each ar In 색범위.Areas
        n = n + ar.Cells.Count
    Next ar
    If n > 120000 Then Exit Sub

    ReDim vC(1 To n)
    ReDim vI(1 To n)
    i = 0
    For Each ar In 색범위.Areas
        For Each cel In ar.Cells
            i = i + 1
            vC(i) = cel.Interior.Color
            vI(i) = cel.Interior.ColorIndex
        Next cel
    Next ar

    mDifWb = 결과범위.Worksheet.Parent.Name
    mDifWs = 결과범위.Worksheet.Name
    mDifResAddr = 결과범위.Address
    mDifColAddr = 색범위.Address
    mDifRes = v
    mDifColor = vC
    mDifIndex = vI
    mDifIns = 삽입함
    mDifInsCol = 붙일열
    mDifOn = True
    Exit Sub
Fail:
    mDifOn = False
End Sub

'------------------------------------------------------------
' 같은 행끼리 비교 되돌리기 (Ctrl+Z 로 호출됨)
'------------------------------------------------------------
Public Sub 행비교되돌리기()
    Dim ws As Worksheet
    Dim 결과범위 As Range, 색범위 As Range
    Dim ar As Range, cel As Range
    Dim i As Long
    Dim oldScreen As Boolean

    If Not mDifOn Or Len(mDifResAddr) = 0 Then
        MsgBox "되돌릴 비교가 없습니다.", vbInformation, "같은 행끼리 비교"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False

    Set ws = Application.Workbooks(mDifWb).Worksheets(mDifWs)
    Set 색범위 = ws.Range(mDifColAddr)

    If mDifIns Then
        ws.Columns(mDifInsCol).Delete Shift:=xlToLeft
    Else
        Set 결과범위 = ws.Range(mDifResAddr)
        For i = 1 To 결과범위.Rows.Count
            결과범위.Cells(i, 1).Formula = mDifRes(i)
        Next i
    End If

    i = 0
    For Each ar In 색범위.Areas
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

    Application.StatusBar = "비교 결과를 되돌렸습니다."
    상태표시줄예약
End Sub

'--- Ctrl 로 두 곳을 잡았을 때 : 무엇을 할지 고르는 작은 팝업 ---
Private Sub 두곳메뉴()
    Dim cb As CommandBar
    Dim btn As CommandBarButton

    On Error Resume Next
    Application.CommandBars(두곳메뉴이름).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=두곳메뉴이름, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    메뉴제목 cb, "고른 두 곳으로 무엇을 할까요"

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  목록끼리 대조   (서로 있는지 1/0, 결과 4열)"
    btn.OnAction = 접두어() & "두곳_목록대조"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "▶  같은 행끼리 비교   (다른 행만 색칠, 결과 1열)"
    btn.OnAction = 접두어() & "두곳_행비교"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     목록끼리 : 값이 상대 열 어딘가에 있는지"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     같은 행끼리 : 같은 줄의 두 칸을 맞대어 봄"
    btn.Style = msoButtonCaption
    btn.Enabled = False

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "메뉴를 표시하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "두 열 비교"
End Sub

Public Sub 두곳_목록대조()
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = Application.ActiveWorkbook.Worksheets(mCmpDirWs)
    대조바로실행 ws.Range(mCmpDirA), ws.Range(mCmpDirB)
    Exit Sub
Fail:
    MsgBox "대조하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "두 열 비교"
End Sub

Public Sub 두곳_행비교()
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = Application.ActiveWorkbook.Worksheets(mCmpDirWs)
    행비교바로 ws.Range(mCmpDirA), ws.Range(mCmpDirB)
    Exit Sub
Fail:
    MsgBox "비교하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "두 열 비교"
End Sub

'------------------------------------------------------------
' 비교 범위 정리
'   열 전체를 찍었을 때 104만 행이 잡히는 것을 막는다.
'   표의 머리글 아래 ~ 마지막 데이터 행까지로 줄인다.
'   두 범위를 같은 구간으로 맞춰 행이 어긋나지 않게 한다.
'------------------------------------------------------------
Private Sub 비교범위정리(ByRef rngA As Range, ByRef rngB As Range)
    Dim ws As Worksheet
    Dim blk As Range
    Dim 마지막 As Range
    Dim 머리글행 As Long, 끝행 As Long
    Dim 넘침 As Boolean

    On Error GoTo Done
    Set ws = rngA.Worksheet
    If rngB.Worksheet.Name <> ws.Name Then Exit Sub

    ' 한쪽이라도 열 전체이거나 20만 행을 넘으면 정리 대상
    넘침 = (rngA.Rows.Count >= ws.Rows.Count) Or (rngB.Rows.Count >= ws.Rows.Count) _
           Or (rngA.Rows.Count > 200000) Or (rngB.Rows.Count > 200000)
    If Not 넘침 Then Exit Sub

    ' A 열에서 값이 있는 마지막 칸을 찾아 그 표를 잡는다
    Set 마지막 = ws.Cells(ws.Rows.Count, rngA.Column).End(xlUp)
    If 마지막.Row < 2 Then
        Set 마지막 = ws.Cells(ws.Rows.Count, rngB.Column).End(xlUp)
        If 마지막.Row < 2 Then Exit Sub
    End If

    Set blk = 소계행제외(마지막.CurrentRegion)
    If blk Is Nothing Then Exit Sub

    머리글행 = blk.Row
    끝행 = blk.Row + blk.Rows.Count - 1
    If 끝행 <= 머리글행 Then Exit Sub

    Set rngA = ws.Range(ws.Cells(머리글행 + 1, rngA.Column), ws.Cells(끝행, rngA.Column))
    Set rngB = ws.Range(ws.Cells(머리글행 + 1, rngB.Column), ws.Cells(끝행, rngB.Column))

Done:
End Sub

'--- 서식 복사 대기 상태 해제 (Esc / 시간 초과 / 붙여넣기 완료 공통) ---
Private Sub 서식복사해제()
    On Error Resume Next
    If mFmtSrcOn Then
        Application.OnKey "{ESC}"
        Application.OnTime mFmtSrcTime, 접두어() & "서식복사시간종료", , False
    End If
    mFmtSrcOn = False
    mFmtSrcAddr = ""
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub 서식복사취소()
    If Not mFmtSrcOn Then
        On Error Resume Next
        Application.OnKey "{ESC}"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    서식복사해제
    Application.StatusBar = "서식 복사를 취소했습니다."
    상태표시줄예약
End Sub

Public Sub 서식복사시간종료()
    If Not mFmtSrcOn Then Exit Sub
    서식복사해제
    Application.StatusBar = "서식 복사 대기가 " & 서식복사_유지초 & "초를 넘겨 자동으로 해제되었습니다."
    상태표시줄예약
End Sub

'--- 열 번호를 열 문자로 (A, B, ... AA) ---
Private Function 열문자(ByVal 열번호 As Long) As String
    Dim s As String
    Dim n As Long
    n = 열번호
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    열문자 = s
End Function

'------------------------------------------------------------
' 텍스트 나누기 설정 되돌리기
'   Excel 은 마지막 TextToColumns 설정을 기억해 두었다가
'   이후 [붙여넣기] 에도 그대로 쓴다.
'   구분 기호를 모두 꺼 둔 채로 두면 탭으로 나뉜 값을 붙여넣어도
'   한 칸에 다 들어가 버린다. 그래서 기본값(탭 구분)으로 돌려 놓는다.
'------------------------------------------------------------
Private Sub 텍스트나누기설정복원()
    Dim ws As Worksheet
    Dim 임시 As Range

    On Error GoTo Done
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Visible = xlSheetVeryHidden

    Set 임시 = ws.Range("A1")
    임시.Value = "a" & vbTab & "b"
    임시.TextToColumns Destination:=임시, _
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
'  쉼표 형식 모양 설정
'    양수      P : N 그냥 · Y + 붙임
'    양수 색   PC: K 검정 · B 파랑 · R 빨강
'    음수      N : T 삼각형 △ · V 역삼각형 ▽ · P 괄호 · M 마이너스
'    음수 색   C : R 빨강 · K 검정 · B 파랑
'    0         Z : D 대시 - · Z 숫자 0 · N 빈칸
'============================================================

Private Function 쉼표양수() As String
    쉼표양수 = 설정읽기("cmaPlus", "N")
End Function
Private Function 쉼표양수색() As String
    쉼표양수색 = 설정읽기("cmaPColor", "K")
End Function
Private Function 쉼표음수() As String
    쉼표음수 = 설정읽기("cmaNeg", "T")
End Function
Private Function 쉼표음수색() As String
    쉼표음수색 = 설정읽기("cmaColor", "R")
End Function
Private Function 쉼표영() As String
    쉼표영 = 설정읽기("cmaZero", "D")
End Function

Private Function 색코드(ByVal c As String) As String
    Select Case c
        Case "B": 색코드 = "[Blue]"
        Case "R": 색코드 = "[Red]"
        Case Else: 색코드 = ""
    End Select
End Function

Private Function 색이름(ByVal c As String) As String
    Select Case c
        Case "B": 색이름 = "파랑"
        Case "R": 색이름 = "빨강"
        Case Else: 색이름 = "검정"
    End Select
End Function

'--- 지금 설정으로 실제 표시 형식 문자열 ---
Private Function 쉼표형식문자열() As String
    Dim 양 As String, 음 As String, 영 As String

    양 = 색코드(쉼표양수색()) & IIf(쉼표양수() = "Y", "+", "") & "#,##0"

    Select Case 쉼표음수()
        Case "V": 음 = 색코드(쉼표음수색()) & "▽#,##0"
        Case "P": 음 = 색코드(쉼표음수색()) & "(#,##0)"
        Case "M": 음 = 색코드(쉼표음수색()) & "-#,##0"
        Case Else: 음 = 색코드(쉼표음수색()) & "△#,##0"
    End Select

    Select Case 쉼표영()
        Case "Z": 영 = "0"
        Case "N": 영 = """"""
        Case Else: 영 = "-"
    End Select

    쉼표형식문자열 = 양 & ";" & 음 & ";" & 영
End Function

'--- 눈에 보이는 예시 ---
Private Function 양수보기() As String
    양수보기 = IIf(쉼표양수() = "Y", "+3,333", "3,333")
End Function

Private Function 음수보기() As String
    Select Case 쉼표음수()
        Case "V": 음수보기 = "▽3,333"
        Case "P": 음수보기 = "(3,333)"
        Case "M": 음수보기 = "-3,333"
        Case Else: 음수보기 = "△3,333"
    End Select
End Function

Private Function 영보기() As String
    Select Case 쉼표영()
        Case "Z": 영보기 = "0"
        Case "N": 영보기 = "빈칸"
        Case Else: 영보기 = "-"
    End Select
End Function

Private Function 쉼표형식한줄() As String
    쉼표형식한줄 = 양수보기() & "  /  " & 음수보기() & "  /  " & 영보기()
End Function

Private Function 음수이름(ByVal c As String) As String
    Select Case c
        Case "V": 음수이름 = "역삼각형 ▽"
        Case "P": 음수이름 = "괄호 ( )"
        Case "M": 음수이름 = "마이너스 -"
        Case Else: 음수이름 = "삼각형 △"
    End Select
End Function

Private Function 영이름(ByVal c As String) As String
    Select Case c
        Case "Z": 영이름 = "숫자 0"
        Case "N": 영이름 = "빈칸"
        Case Else: 영이름 = "대시 -"
    End Select
End Function

Private Function 다음색(ByVal c As String) As String
    Select Case c
        Case "K": 다음색 = "B"
        Case "B": 다음색 = "R"
        Case Else: 다음색 = "K"
    End Select
End Function
Private Function 다음음수(ByVal c As String) As String
    Select Case c
        Case "T": 다음음수 = "V"
        Case "V": 다음음수 = "P"
        Case "P": 다음음수 = "M"
        Case Else: 다음음수 = "T"
    End Select
End Function
Private Function 다음영(ByVal c As String) As String
    Select Case c
        Case "D": 다음영 = "Z"
        Case "Z": 다음영 = "N"
        Case Else: 다음영 = "D"
    End Select
End Function

'------------------------------------------------------------
' 쉼표 형식 모양 - 하위 팝업
'------------------------------------------------------------
Public Sub 쉼표형식메뉴()
    Dim cb As CommandBar
    Dim btn As CommandBarButton

    On Error Resume Next
    Application.CommandBars(쉼표메뉴이름).Delete
    Err.Clear
    On Error GoTo Fail

    Set cb = Application.CommandBars.Add(Name:=쉼표메뉴이름, _
                                         Position:=msoBarPopup, _
                                         MenuBar:=False, Temporary:=True)

    메뉴제목 cb, "쉼표 형식 모양  (Ctrl+Shift+X 로 적용)"

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "     지금 이렇게 나옵니다"
    btn.Style = msoButtonCaption
    btn.Enabled = False
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "          양수     " & 양수보기() & "   (" & 색이름(쉼표양수색()) & ")"
    btn.Style = msoButtonCaption
    btn.Enabled = False

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "          음수     " & 음수보기() & "   (" & 색이름(쉼표음수색()) & ")"
    btn.Style = msoButtonCaption
    btn.Enabled = False

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "          0        " & 영보기()
    btn.Style = msoButtonCaption
    btn.Enabled = False

    '--- 바꾸는 줄 ---
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  양수       " & IIf(쉼표양수() = "Y", "+ 붙임", "그냥") & _
                  "   ->   " & IIf(쉼표양수() = "Y", "그냥", "+ 붙임")
    btn.OnAction = 접두어() & "쉼표양수전환"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  양수 색    " & 색이름(쉼표양수색()) & "   ->   " & 색이름(다음색(쉼표양수색()))
    btn.OnAction = 접두어() & "쉼표양수색전환"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  음수       " & 음수이름(쉼표음수()) & "   ->   " & 음수이름(다음음수(쉼표음수()))
    btn.OnAction = 접두어() & "쉼표음수전환"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  음수 색    " & 색이름(쉼표음수색()) & "   ->   " & 색이름(다음색(쉼표음수색()))
    btn.OnAction = 접두어() & "쉼표음수색전환"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  0          " & 영이름(쉼표영()) & "   ->   " & 영이름(다음영(쉼표영()))
    btn.OnAction = 접두어() & "쉼표영전환"
    btn.Style = msoButtonCaption

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "  기본값으로   (검정 3,333 / 빨강 △3,333 / -)"
    btn.OnAction = 접두어() & "쉼표기본값"
    btn.Style = msoButtonCaption
    btn.BeginGroup = True

    cb.ShowPopup
    Exit Sub
Fail:
    MsgBox "메뉴를 표시하지 못했습니다." & vbCrLf & _
           "오류 " & Err.Number & ": " & Err.Description, vbExclamation, "쉼표 형식"
End Sub

Public Sub 쉼표양수전환()
    설정쓰기 "cmaPlus", IIf(쉼표양수() = "Y", "N", "Y")
    쉼표형식메뉴
End Sub
Public Sub 쉼표양수색전환()
    설정쓰기 "cmaPColor", 다음색(쉼표양수색())
    쉼표형식메뉴
End Sub
Public Sub 쉼표음수전환()
    설정쓰기 "cmaNeg", 다음음수(쉼표음수())
    쉼표형식메뉴
End Sub
Public Sub 쉼표음수색전환()
    설정쓰기 "cmaColor", 다음색(쉼표음수색())
    쉼표형식메뉴
End Sub
Public Sub 쉼표영전환()
    설정쓰기 "cmaZero", 다음영(쉼표영())
    쉼표형식메뉴
End Sub
Public Sub 쉼표기본값()
    설정쓰기 "cmaPlus", "N"
    설정쓰기 "cmaPColor", "K"
    설정쓰기 "cmaNeg", "T"
    설정쓰기 "cmaColor", "R"
    설정쓰기 "cmaZero", "D"
    쉼표형식메뉴
End Sub
