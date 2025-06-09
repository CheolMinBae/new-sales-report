' ========================================================
' 재무 리포트 VBA 빠른 설정 및 보안 해제
' ========================================================

' 이 코드를 먼저 실행하여 환경을 설정하세요
Sub 빠른설정_실행()
    On Error Resume Next
    
    ' 1. 보안 설정 안내
    Call 보안설정_안내
    
    ' 2. 기본 워크시트 설정
    Call 워크시트_기본설정
    
    ' 3. 재무 데이터 템플릿 생성
    Call 재무데이터_템플릿생성
    
    ' 4. 버튼 생성
    Call 버튼_자동생성
    
    ' 4. API 연결 테스트
    Call API연결_확인
    
    MsgBox "설정이 완료되었습니다!" & vbCrLf & _
           "이제 승인/반려 기능을 사용할 수 있습니다.", vbInformation, "설정 완료"
End Sub

' 보안 설정 안내 메시지
Sub 보안설정_안내()
    Dim msg As String
    msg = "매크로 보안 설정 안내:" & vbCrLf & vbCrLf
    msg = msg & "1. 파일 > 옵션 > 보안 센터" & vbCrLf
    msg = msg & "2. 보안 센터 설정 > 매크로 설정" & vbCrLf
    msg = msg & "3. 'VBA 매크로에 대한 알림 표시' 선택" & vbCrLf & vbCrLf
    msg = msg & "또는" & vbCrLf & vbCrLf
    msg = msg & "신뢰할 수 있는 위치에 현재 폴더 추가:" & vbCrLf
    msg = msg & Application.ActiveWorkbook.Path & vbCrLf & vbCrLf
    msg = msg & "이 안내를 보신 후 확인을 눌러주세요."
    
    MsgBox msg, vbInformation, "보안 설정 안내"
End Sub

' 워크시트 기본 설정
Sub 워크시트_기본설정()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 기본 레이블 설정
    ws.Range("A1").Value = "년도:"
    ws.Range("A2").Value = "월:"
    ws.Range("A3").Value = "승인상태:"
    
    ' 기본값 설정
    ws.Range("B1").Value = Year(Date)
    ws.Range("B2").Value = Month(Date)
    ws.Range("D2").Value = "확인 중..."
    
    ' 셀 서식 설정
    ws.Range("A1:A3").Font.Bold = True
    ws.Range("B1:B2").HorizontalAlignment = xlCenter
    ws.Range("D2").HorizontalAlignment = xlCenter
    
    ' 셀 크기 조정
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 15
End Sub

' 버튼 자동 생성
Sub 버튼_자동생성()
    Dim ws As Worksheet
    Dim btnSubmit As Button
    Dim btnPreview As Button
    Dim btnApprove As Button
    Dim btnReject As Button
    Dim btnRefresh As Button
    
    Set ws = ActiveSheet
    
    ' 기존 버튼 삭제
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    ' 데이터 전송 버튼 (가장 중요한 버튼)
    Set btnSubmit = ws.Buttons.Add(150, 50, 80, 25)
    btnSubmit.OnAction = "데이터전송"
    btnSubmit.Caption = "📤 데이터전송"
    btnSubmit.Font.Size = 9
    btnSubmit.Font.Bold = True
    
    ' 미리보기 버튼
    Set btnPreview = ws.Buttons.Add(240, 50, 70, 25)
    btnPreview.OnAction = "데이터전송_미리보기"
    btnPreview.Caption = "👁 미리보기"
    btnPreview.Font.Size = 9
    
    ' 승인 버튼
    Set btnApprove = ws.Buttons.Add(150, 80, 70, 25)
    btnApprove.OnAction = "승인처리"
    btnApprove.Caption = "✅ 승인"
    btnApprove.Font.Size = 10
    btnApprove.Font.Bold = True
    
    ' 반려 버튼
    Set btnReject = ws.Buttons.Add(230, 80, 70, 25)
    btnReject.OnAction = "반려처리"
    btnReject.Caption = "❌ 반려"
    btnReject.Font.Size = 10
    btnReject.Font.Bold = True
    
    ' 새로고침 버튼
    Set btnRefresh = ws.Buttons.Add(310, 80, 70, 25)
    btnRefresh.OnAction = "상태새로고침"
    btnRefresh.Caption = "🔄 새로고침"
    btnRefresh.Font.Size = 9
End Sub

' API 연결 확인
Sub API연결_확인()
    Dim result As String
    result = "API 서버 연결을 확인하는 중..."
    Range("D2").Value = result
    
    ' 실제 API 테스트는 메인 VBA 코드가 있을 때만 가능
    On Error Resume Next
    Call API연결테스트
    On Error GoTo 0
End Sub

' 매크로 보안 상태 확인
Function 매크로보안_확인() As String
    On Error GoTo SecurityError
    
    ' VBA 프로젝트에 접근 시도
    Dim proj As Object
    Set proj = Application.VBE.ActiveVBProject
    
    매크로보안_확인 = "매크로 실행 가능"
    Exit Function
    
SecurityError:
    매크로보안_확인 = "매크로 보안 설정 필요"
End Function

' 파일 저장 안내
Sub 파일저장_안내()
    Dim msg As String
    msg = "중요: 매크로 기능을 유지하려면" & vbCrLf & vbCrLf
    msg = msg & "파일을 저장할 때 반드시" & vbCrLf
    msg = msg & "'Excel 매크로 사용 통합 문서 (*.xlsm)'" & vbCrLf
    msg = msg & "형식으로 저장하세요!" & vbCrLf & vbCrLf
    msg = msg & "Ctrl+S → 파일 형식 → .xlsm 선택"
    
    MsgBox msg, vbExclamation, "파일 저장 안내"
End Sub

' 문제 해결 도움말
Sub 문제해결_도움말()
    Dim msg As String
    msg = "매크로 실행 문제 해결 방법:" & vbCrLf & vbCrLf
    msg = msg & "1. 보안 경고 나타날 때:" & vbCrLf
    msg = msg & "   → '콘텐츠 사용' 클릭" & vbCrLf & vbCrLf
    msg = msg & "2. 매크로 차단될 때:" & vbCrLf
    msg = msg & "   → 파일 > 옵션 > 보안 센터" & vbCrLf
    msg = msg & "   → 매크로 설정 변경" & vbCrLf & vbCrLf
    msg = msg & "3. 신뢰할 수 있는 위치:" & vbCrLf
    msg = msg & "   → " & Application.ActiveWorkbook.Path & vbCrLf & vbCrLf
    msg = msg & "4. 관리자 권한으로 Excel 실행"
    
    MsgBox msg, vbInformation, "문제 해결 가이드"
End Sub 