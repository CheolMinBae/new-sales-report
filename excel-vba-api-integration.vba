' ========================================================
' 재무 리포트 대시보드 API 연동 VBA 코드
' ========================================================

Option Explicit

' API 기본 설정
Private Const API_BASE_URL As String = "http://localhost:3001/api"
Private Const EXCEL_VERSION As String = "Excel VBA v1.0"

' ===== 메인 버튼 이벤트 =====

' 승인 버튼 클릭 시
Sub 승인처리()
    Dim month As Integer
    Dim year As Integer
    Dim memo As String
    
    ' 현재 월/년도 가져오기 (셀에서 읽거나 기본값 사용)
    month = GetCurrentMonth()
    year = GetCurrentYear()
    
    ' 메모 입력받기
    memo = InputBox("승인 메모를 입력하세요 (선택사항):", "승인 처리", "")
    
    ' 승인 처리 실행
    If SendApprovalToAPI(month, year, "approved", memo) Then
        MsgBox year & "년 " & month & "월 레포트가 성공적으로 승인되었습니다.", vbInformation, "승인 완료"
        RefreshApprovalStatus
    Else
        MsgBox "승인 처리 중 오류가 발생했습니다.", vbCritical, "오류"
    End If
End Sub

' 반려 버튼 클릭 시
Sub 반려처리()
    Dim month As Integer
    Dim year As Integer
    Dim memo As String
    
    month = GetCurrentMonth()
    year = GetCurrentYear()
    
    ' 반려 사유 입력받기 (필수)
    memo = InputBox("반려 사유를 입력하세요:", "반려 처리", "")
    If memo = "" Then
        MsgBox "반려 사유는 필수입니다.", vbExclamation, "입력 필요"
        Exit Sub
    End If
    
    ' 반려 처리 실행
    If SendApprovalToAPI(month, year, "rejected", memo) Then
        MsgBox year & "년 " & month & "월 레포트가 반려되었습니다.", vbInformation, "반려 완료"
        RefreshApprovalStatus
    Else
        MsgBox "반려 처리 중 오류가 발생했습니다.", vbCritical, "오류"
    End If
End Sub

' 상태 새로고침 버튼
Sub 상태새로고침()
    RefreshApprovalStatus
    MsgBox "승인 상태가 새로고침되었습니다.", vbInformation, "새로고침 완료"
End Sub

' ===== API 통신 함수 =====

' API로 승인/반려 정보 전송
Function SendApprovalToAPI(month As Integer, year As Integer, approvalStatus As String, memo As String) As Boolean
    Dim http As Object
    Dim url As String
    Dim jsonData As String
    Dim response As String
    
    On Error GoTo ErrorHandler
    
    ' WinHTTP 객체 생성
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' API URL 설정
    url = API_BASE_URL & "/excel"
    
    ' JSON 데이터 생성
    jsonData = "{"
    jsonData = jsonData & """month"": " & month & ","
    jsonData = jsonData & """year"": " & year & ","
    jsonData = jsonData & """approvalStatus"": """ & approvalStatus & ""","
    jsonData = jsonData & """memo"": """ & EscapeJsonString(memo) & ""","
    jsonData = jsonData & """approvedBy"": """ & Application.UserName & ""","
    jsonData = jsonData & """excelVersion"": """ & EXCEL_VERSION & """"
    jsonData = jsonData & "}"
    
    ' HTTP 요청 설정
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    
    ' 요청 전송
    http.Send jsonData
    
    ' 응답 확인
    If http.Status = 200 Then
        response = http.ResponseText
        ' JSON 응답에서 success 필드 확인
        If InStr(response, """success"":true") > 0 Then
            SendApprovalToAPI = True
        Else
            SendApprovalToAPI = False
        End If
    Else
        SendApprovalToAPI = False
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    SendApprovalToAPI = False
    Set http = Nothing
End Function

' API에서 승인 상태 조회
Function GetApprovalStatusFromAPI(month As Integer, year As Integer) As String
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim status As String
    
    On Error GoTo ErrorHandler
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' API URL 설정
    url = API_BASE_URL & "/excel?month=" & month & "&year=" & year
    
    ' HTTP GET 요청
    http.Open "GET", url, False
    http.Send
    
    If http.Status = 200 Then
        response = http.ResponseText
        ' JSON에서 approvalStatus 추출 (간단한 파싱)
        status = ExtractJsonValue(response, "approvalStatus")
        GetApprovalStatusFromAPI = status
    Else
        GetApprovalStatusFromAPI = "error"
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    GetApprovalStatusFromAPI = "error"
    Set http = Nothing
End Function

' ===== 유틸리티 함수 =====

' 현재 월 가져오기 (셀 B2에서 읽거나 현재 월 사용)
Function GetCurrentMonth() As Integer
    Dim cellValue As Variant
    cellValue = Range("B2").Value
    
    If IsNumeric(cellValue) And cellValue >= 1 And cellValue <= 12 Then
        GetCurrentMonth = CInt(cellValue)
    Else
        GetCurrentMonth = Month(Date)
    End If
End Function

' 현재 년도 가져오기 (셀 B1에서 읽거나 현재 년도 사용)
Function GetCurrentYear() As Integer
    Dim cellValue As Variant
    cellValue = Range("B1").Value
    
    If IsNumeric(cellValue) And cellValue >= 2020 And cellValue <= 2030 Then
        GetCurrentYear = CInt(cellValue)
    Else
        GetCurrentYear = Year(Date)
    End If
End Function

' JSON 문자열 이스케이프 처리
Function EscapeJsonString(inputStr As String) As String
    Dim result As String
    result = inputStr
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    EscapeJsonString = result
End Function

' JSON에서 값 추출 (간단한 파싱)
Function ExtractJsonValue(jsonStr As String, key As String) As String
    Dim searchStr As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String
    
    searchStr = """" & key & """:"
    startPos = InStr(jsonStr, searchStr)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        ' 값의 시작 위치 찾기 (따옴표 다음)
        startPos = InStr(startPos, jsonStr, """") + 1
        ' 값의 끝 위치 찾기
        endPos = InStr(startPos, jsonStr, """")
        
        If endPos > startPos Then
            value = Mid(jsonStr, startPos, endPos - startPos)
            ExtractJsonValue = value
        Else
            ExtractJsonValue = ""
        End If
    Else
        ExtractJsonValue = ""
    End If
End Function

' 승인 상태 새로고침 및 셀 업데이트
Sub RefreshApprovalStatus()
    Dim month As Integer
    Dim year As Integer
    Dim status As String
    Dim statusText As String
    
    month = GetCurrentMonth()
    year = GetCurrentYear()
    
    status = GetApprovalStatusFromAPI(month, year)
    
    ' 상태를 한국어로 변환
    Select Case status
        Case "approved"
            statusText = "승인완료"
        Case "rejected"
            statusText = "반려"
        Case "pending"
            statusText = "승인대기"
        Case Else
            statusText = "확인불가"
    End Select
    
    ' 상태를 셀에 표시 (예: D2 셀)
    Range("D2").Value = statusText
    
    ' 상태에 따라 셀 색상 변경
    Select Case status
        Case "approved"
            Range("D2").Interior.Color = RGB(144, 238, 144) ' 연한 녹색
        Case "rejected"
            Range("D2").Interior.Color = RGB(255, 182, 193) ' 연한 빨강
        Case "pending"
            Range("D2").Interior.Color = RGB(255, 255, 224) ' 연한 노랑
        Case Else
            Range("D2").Interior.Color = RGB(211, 211, 211) ' 회색
    End Select
End Sub

' ===== 자동 실행 함수 =====

' 워크북 열릴 때 자동으로 상태 새로고침
Sub Auto_Open()
    RefreshApprovalStatus
End Sub

' 워크북이 활성화될 때 자동으로 상태 새로고침
Sub Workbook_Activate()
    RefreshApprovalStatus
End Sub

' ===== 설정 및 초기화 =====

' 버튼 및 UI 설정 (한 번만 실행)
Sub 버튼설정()
    Dim ws As Worksheet
    Dim btnApprove As Button
    Dim btnReject As Button
    Dim btnRefresh As Button
    
    Set ws = ActiveSheet
    
    ' 기존 버튼 삭제
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    ' 승인 버튼 생성
    Set btnApprove = ws.Buttons.Add(200, 50, 80, 30)
    btnApprove.OnAction = "승인처리"
    btnApprove.Caption = "승인"
    
    ' 반려 버튼 생성
    Set btnReject = ws.Buttons.Add(290, 50, 80, 30)
    btnReject.OnAction = "반려처리"
    btnReject.Caption = "반려"
    
    ' 새로고침 버튼 생성
    Set btnRefresh = ws.Buttons.Add(380, 50, 80, 30)
    btnRefresh.OnAction = "상태새로고침"
    btnRefresh.Caption = "새로고침"
    
    ' 라벨 설정
    Range("A1").Value = "년도:"
    Range("A2").Value = "월:"
    Range("A3").Value = "승인상태:"
    
    ' 기본값 설정
    Range("B1").Value = Year(Date)
    Range("B2").Value = Month(Date)
    
    MsgBox "버튼 설정이 완료되었습니다.", vbInformation, "설정 완료"
End Sub

' API 연결 테스트
Sub API연결테스트()
    Dim month As Integer
    Dim year As Integer
    Dim status As String
    
    month = GetCurrentMonth()
    year = GetCurrentYear()
    
    status = GetApprovalStatusFromAPI(month, year)
    
    If status <> "error" Then
        MsgBox "API 연결 성공!" & vbCrLf & year & "년 " & month & "월 상태: " & status, vbInformation, "연결 테스트"
    Else
        MsgBox "API 연결 실패!" & vbCrLf & "서버가 실행 중인지 확인하세요.", vbCritical, "연결 오류"
    End If
End Sub 