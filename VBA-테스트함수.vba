' ========================================================
' API 연결 테스트용 VBA 함수들
' ========================================================

' 간단한 API 연결 테스트
Sub 간단한_API_테스트()
    Dim result As String
    result = CallTestAPI()
    
    If result <> "error" Then
        MsgBox "API 연결 성공!" & vbCrLf & vbCrLf & result, vbInformation, "연결 테스트 성공"
    Else
        MsgBox "API 연결 실패!" & vbCrLf & "서버가 실행 중인지 확인하세요.", vbCritical, "연결 오류"
    End If
End Sub

' 테스트 API 호출 함수
Function CallTestAPI() As String
    Dim http As Object
    Dim url As String
    Dim response As String
    
    On Error GoTo ErrorHandler
    
    ' WinHTTP 객체 생성
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' API URL 설정 (메시지 파라미터 포함)
    url = "http://sales-report-alb-848109300.ap-northeast-2.elb.amazonaws.com/api/test?message=VBA에서 안녕하세요!"
    
    ' HTTP GET 요청
    http.Open "GET", url, False
    http.Send
    
    ' 응답 확인
    If http.Status = 200 Then
        response = http.ResponseText
        CallTestAPI = response
    Else
        CallTestAPI = "error"
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    CallTestAPI = "error"
    Set http = Nothing
End Function

' JSON에서 특정 값 추출하여 표시
Sub API응답_상세보기()
    Dim response As String
    Dim success As String
    Dim message As String
    Dim timestamp As String
    
    response = CallTestAPI()
    
    If response <> "error" Then
        ' JSON에서 값들 추출
        success = ExtractJsonValue(response, "success")
        message = ExtractJsonValue(response, "message")
        timestamp = ExtractJsonValue(response, "timestamp")
        
        ' 결과를 셀에 표시
        Range("F1").Value = "API 테스트 결과:"
        Range("F2").Value = "성공 여부: " & success
        Range("F3").Value = "메시지: " & message
        Range("F4").Value = "시간: " & timestamp
        
        ' 셀 서식 설정
        Range("F1").Font.Bold = True
        Range("F1:F4").Font.Size = 10
        
        MsgBox "API 응답 상세 정보가 F열에 표시되었습니다.", vbInformation, "상세 정보"
    Else
        MsgBox "API 호출에 실패했습니다.", vbCritical, "오류"
    End If
End Sub

' URL과 포트 연결 테스트
Sub 포트연결_테스트()
    Dim ports As Variant
    Dim i As Integer
    Dim result As String
    
    ' 테스트할 포트들
    ports = Array(3000, 3001, 8080, 5000)
    
    For i = 0 To UBound(ports)
        result = TestPortConnection(ports(i))
        Debug.Print "포트 " & ports(i) & ": " & result
    Next i
    
    MsgBox "포트 연결 테스트 완료!" & vbCrLf & "결과는 직접 실행 창(Ctrl+G)에서 확인하세요.", vbInformation, "포트 테스트"
End Sub

' 특정 포트 연결 테스트
Function TestPortConnection(port As Integer) As String
    Dim http As Object
    Dim url As String
    
    On Error GoTo ErrorHandler
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://sales-report-alb-848109300.ap-northeast-2.elb.amazonaws.com/api/test"
    
    http.Open "GET", url, False
    http.SetTimeouts 1000, 1000, 1000, 1000  ' 1초 타임아웃
    http.Send
    
    If http.Status = 200 Then
        TestPortConnection = "연결 성공"
    Else
        TestPortConnection = "HTTP " & http.Status
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    TestPortConnection = "연결 실패"
    Set http = Nothing
End Function

' 현재 서버 상태 확인
Sub 서버상태_확인()
    Dim ws As Worksheet
    Dim lastRow As Integer
    
    Set ws = ActiveSheet
    
    ' 헤더 추가
    ws.Range("H1").Value = "서버 상태 확인"
    ws.Range("H2").Value = "시간"
    ws.Range("I2").Value = "상태"
    ws.Range("J2").Value = "응답시간"
    
    ' 헤더 서식
    ws.Range("H1:J2").Font.Bold = True
    ws.Range("H2:J2").Interior.Color = RGB(200, 200, 200)
    
    lastRow = 3
    
    ' 현재 시간과 상태 기록
    ws.Range("H" & lastRow).Value = Now()
    
    Dim startTime As Double
    Dim endTime As Double
    Dim response As String
    
    startTime = Timer
    response = CallTestAPI()
    endTime = Timer
    
    If response <> "error" Then
        ws.Range("I" & lastRow).Value = "정상"
        ws.Range("I" & lastRow).Interior.Color = RGB(144, 238, 144)  ' 연한 녹색
    Else
        ws.Range("I" & lastRow).Value = "오류"
        ws.Range("I" & lastRow).Interior.Color = RGB(255, 182, 193)  ' 연한 빨강
    End If
    
    ws.Range("J" & lastRow).Value = Format((endTime - startTime), "0.000") & "초"
    
    MsgBox "서버 상태가 H열에 기록되었습니다.", vbInformation, "상태 확인 완료"
End Sub 