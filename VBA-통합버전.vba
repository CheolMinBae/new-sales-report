' ========================================================
' 재무 리포트 대시보드 API 연동 VBA 통합 버전
' 모든 필요한 함수들이 포함된 완전 버전
' ========================================================

Option Explicit

' 재무 데이터 구조체 (Dictionary 대신 사용)
Type FinanceData
    salesRevenue As Double     ' 매출
    otherIncome As Double      ' 기타수입  
    rentExpense As Double      ' 임대료
    laborExpense As Double     ' 인건비
    materialExpense As Double  ' 재료비
    operatingExpense As Double ' 운영비
    otherExpense As Double     ' 기타비용
    cashBalance As Double      ' 현금잔고
End Type

' API 기본 설정
Private Const API_BASE_URL As String = "http://localhost:3001/api"
Private Const EXCEL_VERSION As String = "Excel VBA v1.0"

' ===== 메인 버튼 이벤트 =====

' 데이터 전송 버튼 클릭 시
Sub 데이터전송()
    Dim year As Integer
    Dim month As Integer
    Dim result As Boolean
    
    year = GetCurrentYear()
    month = GetCurrentMonth()
    
    ' 데이터 유효성 검사
    If Not ValidateFinanceData() Then
        MsgBox "재무 데이터를 확인해주세요. 필수 항목이 누락되었습니다.", vbExclamation, "데이터 확인 필요"
        Exit Sub
    End If
    
    ' 데이터 전송 실행 (SendFinanceDataToAPI 함수 내에서 확인 메시지와 응답 표시)
    result = SendFinanceDataToAPI(year, month)
    
    ' 전송 성공 시 상태 새로고침
    If result Then
        RefreshApprovalStatus
    End If
End Sub

' 데이터 전송 전 미리보기 (디버깅 정보 포함)
Sub 데이터전송_미리보기()
    Dim year As Integer
    Dim month As Integer
    Dim financeData As FinanceData
    Dim msg As String
    
    year = GetCurrentYear()
    month = GetCurrentMonth()
    financeData = ReadFinanceDataFromCells()
    
    ' 미리보기 메시지 구성
    msg = year & "년 " & month & "월 재무 데이터 미리보기:" & vbCrLf & vbCrLf
    msg = msg & "매출: " & Format(financeData.salesRevenue, "#,##0") & "원" & vbCrLf
    msg = msg & "기타수입: " & Format(financeData.otherIncome, "#,##0") & "원" & vbCrLf
    msg = msg & "총 매출: " & Format(financeData.salesRevenue + financeData.otherIncome, "#,##0") & "원" & vbCrLf & vbCrLf
    msg = msg & "임대료: " & Format(financeData.rentExpense, "#,##0") & "원" & vbCrLf
    msg = msg & "인건비: " & Format(financeData.laborExpense, "#,##0") & "원" & vbCrLf
    msg = msg & "재료비: " & Format(financeData.materialExpense, "#,##0") & "원" & vbCrLf
    msg = msg & "운영비: " & Format(financeData.operatingExpense, "#,##0") & "원" & vbCrLf
    msg = msg & "기타비용: " & Format(financeData.otherExpense, "#,##0") & "원" & vbCrLf
    msg = msg & "총 지출: " & Format(financeData.rentExpense + financeData.laborExpense + financeData.materialExpense + financeData.operatingExpense + financeData.otherExpense, "#,##0") & "원" & vbCrLf & vbCrLf
    msg = msg & "현금잔고: " & Format(financeData.cashBalance, "#,##0") & "원" & vbCrLf & vbCrLf
    msg = msg & "이 데이터를 전송하시겠습니까?"
    
    If MsgBox(msg, vbQuestion + vbYesNo, "데이터 전송 미리보기") = vbYes Then
        Call 데이터전송
    End If
End Sub

' 데이터 수집 디버깅 - 어떤 시트에서 데이터를 찾았는지 확인
Sub 데이터수집_디버깅()
    Dim year As Integer
    Dim month As Integer
    Dim ws As Worksheet
    Dim debugMsg As String
    Dim salesFromTable As Double
    Dim salesFromBank As Double
    Dim otherIncome As Double
    
    year = GetCurrentYear()
    month = GetCurrentMonth()
    
    debugMsg = "🔍 " & year & "년 " & month & "월 데이터 수집 디버깅:" & vbCrLf & vbCrLf
    
    ' === 시트 존재 확인 ===
    debugMsg = debugMsg & "📋 시트 존재 확인:" & vbCrLf
    
    ' 1번 시트 (정리표) 확인
    On Error Resume Next
    Set ws = Worksheets(1) ' 1번 시트 = 정리표
    On Error GoTo 0
    
    If ws Is Nothing Then
        debugMsg = debugMsg & "❌ 1번 시트 (정리표) 없음" & vbCrLf
    Else
        debugMsg = debugMsg & "✅ 1번 시트 (정리표): " & ws.Name & vbCrLf
        salesFromTable = FindMonthlyDataInSheet(ws, year, month, "매출입금", "매출")
        otherIncome = FindMonthlyDataInSheet(ws, year, month, "기타입금", "기타")
        debugMsg = debugMsg & "   매출: " & Format(salesFromTable, "#,##0") & "원" & vbCrLf
        debugMsg = debugMsg & "   기타수입: " & Format(otherIncome, "#,##0") & "원" & vbCrLf
    End If
    
    ' 2번 시트 (통장) 확인
    On Error Resume Next
    Set ws = Nothing
    Set ws = Worksheets(2) ' 2번 시트 = 통장
    On Error GoTo 0
    
    If ws Is Nothing Then
        debugMsg = debugMsg & "❌ 2번 시트 (통장) 없음" & vbCrLf
    Else
        debugMsg = debugMsg & "✅ 2번 시트 (통장): " & ws.Name & vbCrLf
        salesFromBank = SumMonthlyTransactions(ws, year, month, "매출입금")
        debugMsg = debugMsg & "   매출입금 합계: " & Format(salesFromBank, "#,##0") & "원" & vbCrLf
    End If
    
    ' 3번 시트 (캐시플로우) 확인
    On Error Resume Next
    Set ws = Nothing
    Set ws = Worksheets(3) ' 3번 시트 = 캐시플로우
    On Error GoTo 0
    
    If ws Is Nothing Then
        debugMsg = debugMsg & "❌ 3번 시트 (캐시플로우) 없음" & vbCrLf
    Else
        debugMsg = debugMsg & "✅ 3번 시트 (캐시플로우): " & ws.Name & vbCrLf
    End If
    
    debugMsg = debugMsg & vbCrLf
    debugMsg = debugMsg & "📊 최종 합계:" & vbCrLf
    debugMsg = debugMsg & "총 매출: " & Format(salesFromTable + salesFromBank, "#,##0") & "원" & vbCrLf
    debugMsg = debugMsg & "(정리표: " & Format(salesFromTable, "#,##0") & " + 통장: " & Format(salesFromBank, "#,##0") & ")"
    
    MsgBox debugMsg, vbInformation, "데이터 수집 디버깅"
End Sub

' 시트 구조 분석 - 실제 시트의 구조를 확인
Sub 시트구조_분석()
    Dim ws As Worksheet
    Dim msg As String
    Dim i As Long, j As Long
    Dim year As Integer
    
    year = GetCurrentYear()
    
    ' 사용자가 분석할 시트 선택 (시트 순서 기준)
    Dim sheetNumber As String
    sheetNumber = InputBox("분석할 시트 번호를 입력하세요:" & vbCrLf & vbCrLf & _
                        "1번: 정리표 (20~25년 정리표)" & vbCrLf & _
                        "2번: 통장 (2020년-통장)" & vbCrLf & _
                        "3번: 캐시플로우 (CASH FLOW-2020년)", "시트 구조 분석", "1")
    
    If sheetNumber = "" Then Exit Sub
    
    Dim sheetIndex As Integer
    sheetIndex = Val(sheetNumber)
    
    If sheetIndex < 1 Or sheetIndex > 3 Then
        MsgBox "1, 2, 3 중 하나의 번호를 입력하세요.", vbExclamation, "잘못된 입력"
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = Worksheets(sheetIndex)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox sheetIndex & "번 시트를 찾을 수 없습니다.", vbExclamation, "시트 없음"
        Exit Sub
    End If
    
    msg = "📋 " & sheetIndex & "번 시트 (" & ws.Name & ") 구조 분석:" & vbCrLf & vbCrLf
    
    ' 처음 10행 x 10열 데이터 표시
    msg = msg & "📊 데이터 미리보기 (10x10):" & vbCrLf
    For i = 1 To 10
        For j = 1 To 10
            If j = 1 Then
                msg = msg & "행" & i & ": "
            End If
            msg = msg & Chr(64 + j) & "=" & Left(ws.Cells(i, j).Value, 8) & " | "
        Next j
        msg = msg & vbCrLf
    Next i
    
    ' 특정 키워드 검색
    msg = msg & vbCrLf & "🔍 키워드 검색 결과:" & vbCrLf
    Dim keywords As Variant
    keywords = Array(year, "매출입금", "기타입금", "비용결제", "외상대", "현금잔고", "1월", "2월", "3월")
    
    For i = LBound(keywords) To UBound(keywords)
        Dim foundCells As String
        foundCells = FindKeywordInSheet(ws, keywords(i))
        If foundCells <> "" Then
            msg = msg & "• " & keywords(i) & ": " & foundCells & vbCrLf
        Else
            msg = msg & "• " & keywords(i) & ": 없음" & vbCrLf
        End If
    Next i
    
    MsgBox msg, vbInformation, "시트 구조 분석"
End Sub

' 시트에서 키워드 찾기
Function FindKeywordInSheet(ws As Worksheet, keyword As Variant) As String
    Dim searchRange As Range
    Dim foundCell As Range
    Dim result As String
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow > 50 Then lastRow = 50 ' 검색 범위 제한
    If lastCol > 20 Then lastCol = 20
    
    Set searchRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    Set foundCell = searchRange.Find(keyword, LookIn:=xlValues, LookAt:=xlPart)
    
    If Not foundCell Is Nothing Then
        result = Chr(64 + foundCell.Column) & foundCell.Row
        ' 추가로 더 찾기
        Dim firstAddress As String
        firstAddress = foundCell.Address
        Do
            Set foundCell = searchRange.FindNext(foundCell)
            If foundCell.Address <> firstAddress And result <> "" Then
                result = result & ", " & Chr(64 + foundCell.Column) & foundCell.Row
            End If
        Loop While foundCell.Address <> firstAddress And Len(result) < 50
    Else
        result = ""
    End If
    
    FindKeywordInSheet = result
End Function

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
    
    ' 승인 처리 실행 (SendApprovalToAPI 함수 내에서 확인 메시지와 응답 표시)
    If SendApprovalToAPI(month, year, "approved", memo) Then
        RefreshApprovalStatus
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
    
    ' 반려 처리 실행 (SendApprovalToAPI 함수 내에서 확인 메시지와 응답 표시)
    If SendApprovalToAPI(month, year, "rejected", memo) Then
        RefreshApprovalStatus
    End If
End Sub

' 상태 새로고침 버튼
Sub 상태새로고침()
    RefreshApprovalStatus
    MsgBox "승인 상태가 새로고침되었습니다.", vbInformation, "새로고침 완료"
End Sub

' 승인상태확인 버튼 - 테이블의 해당 월 row에 승인상태 업데이트
Sub 승인상태확인()
    Dim ws As Worksheet
    Dim year As Integer
    Dim month As Integer
    Dim status As String
    Dim statusText As String
    Dim targetRow As Long
    Dim targetCol As Long
    Dim monthNames As Variant
    Dim i As Integer
    Dim foundRow As Long
    Dim confirmMsg As String
    
    ' 대시보드 시트 찾기
    On Error Resume Next
    Set ws = Worksheets("재무리포트_대시보드")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    ' 데이터 전송용 년도/월 가져오기
    year = GetCurrentYear()
    month = GetCurrentMonth()
    
    ' 확인 메시지
    confirmMsg = "📋 승인상태 확인 및 업데이트" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "대상: " & year & "년 " & month & "월" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "현재 시트의 해당 월 데이터에" & vbCrLf
    confirmMsg = confirmMsg & "승인상태를 업데이트하시겠습니까?" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "⚠️ 주의: 기존 데이터가 있으면 덮어씁니다."
    
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "승인상태 확인") = vbNo Then
        Exit Sub
    End If
    
    ' API에서 승인상태 가져오기
    status = GetApprovalStatusFromAPI(month, year)
    
    ' 상태를 한국어로 변환
    Select Case status
        Case "approved"
            statusText = "승인완료"
        Case "rejected"
            statusText = "반려"
        Case "pending"
            statusText = "승인대기"
        Case "error"
            statusText = "연결오류"
        Case Else
            statusText = "확인불가"
    End Select
    
    ' 월별 데이터 테이블에서 해당 월 찾기 및 업데이트
    Call 월별테이블_승인상태업데이트(ws, year, month, statusText, status)
    
    ' 대시보드의 상태도 업데이트
    Call RefreshApprovalStatus
    
    ' 결과 메시지
    Dim resultMsg As String
    resultMsg = "✅ 승인상태 확인 완료!" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "📅 대상: " & year & "년 " & month & "월" & vbCrLf
    resultMsg = resultMsg & "📊 상태: " & statusText & vbCrLf & vbCrLf
    
    If status <> "error" Then
        resultMsg = resultMsg & "✨ 테이블이 성공적으로 업데이트되었습니다."
    Else
        resultMsg = resultMsg & "⚠️ API 연결 오류가 발생했습니다." & vbCrLf
        resultMsg = resultMsg & "서버 상태를 확인하세요."
    End If
    
    MsgBox resultMsg, vbInformation, "승인상태 확인 완료"
End Sub

' 전체월 승인상태 확인 및 업데이트 (새로고침 버튼용)
Sub 전체월_승인상태확인()
    Dim ws As Worksheet
    Dim year As Integer
    Dim month As Integer
    Dim status As String
    Dim statusText As String
    Dim successCount As Integer
    Dim failCount As Integer
    Dim resultMsg As String
    Dim confirmMsg As String
    Dim monthlyResults As String
    
    ' 대시보드 시트 찾기
    On Error Resume Next
    Set ws = Worksheets("재무리포트_대시보드")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
    ' 승인상태 확인용 년도 가져오기 (B7 셀)
    year = GetApprovalStatusYear()
    
    ' 확인 메시지
    confirmMsg = "🔄 전체월 승인상태 새로고침" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "대상 년도: " & year & "년" & vbCrLf
    confirmMsg = confirmMsg & "확인 범위: 1월 ~ 12월 (전체)" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "A8~A19의 각 월에 승인상태를" & vbCrLf
    confirmMsg = confirmMsg & "업데이트하시겠습니까?" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "⚠️ 주의: API 호출이 12번 발생합니다." & vbCrLf
    confirmMsg = confirmMsg & "네트워크 상태를 확인하세요."
    
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "승인상태 새로고침") = vbNo Then
        Exit Sub
    End If
    
    ' 시작 메시지
    Application.StatusBar = "전체월 승인상태 새로고침 중..."
    monthlyResults = "🔄 " & year & "년 전체월 승인상태 새로고침 결과:" & vbCrLf & vbCrLf
    
    ' 1월부터 12월까지 순차 확인 및 업데이트
    For month = 1 To 12
        Application.StatusBar = "승인상태 확인 중... (" & month & "/12)"
        
        ' API에서 승인상태 가져오기
        status = GetApprovalStatusFromAPI(month, year)
        
        ' 상태를 한국어로 변환
        Select Case status
            Case "approved"
                statusText = "승인완료"
                successCount = successCount + 1
            Case "rejected"
                statusText = "반려"
                successCount = successCount + 1
            Case "pending"
                statusText = "승인대기"
                successCount = successCount + 1
            Case "error"
                statusText = "연결오류"
                failCount = failCount + 1
            Case Else
                statusText = "확인불가"
                failCount = failCount + 1
        End Select
        
        ' A8~A19의 각 월 행에 승인상태 업데이트 (B열에)
        Call 월별리스트_승인상태업데이트(ws, month, statusText, status)
        
        ' 결과 기록
        monthlyResults = monthlyResults & month & "월: " & statusText
        If status = "error" Or status = "" Then
            monthlyResults = monthlyResults & " ❌"
        Else
            monthlyResults = monthlyResults & " ✅"
        End If
        monthlyResults = monthlyResults & vbCrLf
        
        ' 잠시 대기 (API 부하 방지)
        Application.Wait (Now + TimeValue("0:00:01"))
    Next month
    
    ' 상태바 초기화
    Application.StatusBar = False
    
    ' 최종 결과 메시지
    resultMsg = "🎉 전체월 승인상태 새로고침 완료!" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "📅 확인 년도: " & year & "년" & vbCrLf
    resultMsg = resultMsg & "✅ 성공: " & successCount & "개월" & vbCrLf
    resultMsg = resultMsg & "❌ 실패: " & failCount & "개월" & vbCrLf & vbCrLf
    resultMsg = resultMsg & monthlyResults & vbCrLf
    resultMsg = resultMsg & "⏰ 완료 시간: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
    
    ' 결과 메시지 박스 (요약)
    Dim summaryMsg As String
    summaryMsg = "🎉 전체월 승인상태 새로고침 완료!" & vbCrLf & vbCrLf
    summaryMsg = summaryMsg & "📊 결과 요약:" & vbCrLf
    summaryMsg = summaryMsg & "✅ 성공: " & successCount & "개월" & vbCrLf
    summaryMsg = summaryMsg & "❌ 실패: " & failCount & "개월" & vbCrLf & vbCrLf
    summaryMsg = summaryMsg & "📋 A8~A19의 각 월에 승인상태가" & vbCrLf
    summaryMsg = summaryMsg & "업데이트되었습니다."
    
    MsgBox summaryMsg, vbInformation, "승인상태 새로고침 완료"
End Sub

' A8~A19 월별 리스트에 승인상태 업데이트
Sub 월별리스트_승인상태업데이트(ws As Worksheet, month As Integer, statusText As String, status As String)
    Dim targetRow As Long
    Dim statusCol As Long
    
    ' 해당 월의 행 계산 (A8=1월, A9=2월, ..., A19=12월)
    targetRow = 7 + month  ' A8부터 시작하므로 7을 더함
    statusCol = 2  ' B열에 승인상태 기록
    
    ' 승인상태 업데이트
    ws.Cells(targetRow, statusCol).Value = statusText
    
    ' 셀 색상 설정
    Select Case status
        Case "approved"
            ws.Cells(targetRow, statusCol).Interior.Color = RGB(144, 238, 144) ' 연한 녹색
            ws.Cells(targetRow, statusCol).Font.Color = RGB(0, 100, 0)
        Case "rejected"
            ws.Cells(targetRow, statusCol).Interior.Color = RGB(255, 182, 193) ' 연한 빨강
            ws.Cells(targetRow, statusCol).Font.Color = RGB(150, 0, 0)
        Case "pending"
            ws.Cells(targetRow, statusCol).Interior.Color = RGB(255, 255, 224) ' 연한 노랑
            ws.Cells(targetRow, statusCol).Font.Color = RGB(150, 150, 0)
        Case Else
            ws.Cells(targetRow, statusCol).Interior.Color = RGB(211, 211, 211) ' 회색
            ws.Cells(targetRow, statusCol).Font.Color = RGB(100, 100, 100)
    End Select
    
    ' 셀 서식 설정
    ws.Cells(targetRow, statusCol).HorizontalAlignment = xlCenter
    ws.Cells(targetRow, statusCol).Font.Bold = True
    ws.Cells(targetRow, statusCol).Borders.LineStyle = xlContinuous
    ws.Cells(targetRow, statusCol).Font.Size = 10
End Sub

' 월별 테이블에서 승인상태 업데이트
Sub 월별테이블_승인상태업데이트(ws As Worksheet, year As Integer, month As Integer, statusText As String, status As String)
    Dim searchRange As Range
    Dim foundCell As Range
    Dim targetRow As Long
    Dim statusCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim monthText As String
    
    ' 월을 텍스트로 변환 (1월, 2월, ... 형태로 검색)
    monthText = month & "월"
    
    ' 현재 시트에서 마지막 행과 열 찾기
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' 전체 범위에서 해당 월 찾기
    Set searchRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    Set foundCell = searchRange.Find(monthText, LookIn:=xlValues, LookAt:=xlPart)
    
    If Not foundCell Is Nothing Then
        targetRow = foundCell.Row
        
        ' 승인상태 열 찾기 또는 생성
        statusCol = 월별테이블_승인상태열찾기(ws, lastCol)
        
        ' 승인상태 업데이트
        ws.Cells(targetRow, statusCol).Value = statusText
        
        ' 셀 색상 설정
        Select Case status
            Case "approved"
                ws.Cells(targetRow, statusCol).Interior.Color = RGB(144, 238, 144) ' 연한 녹색
                ws.Cells(targetRow, statusCol).Font.Color = RGB(0, 100, 0)
            Case "rejected"
                ws.Cells(targetRow, statusCol).Interior.Color = RGB(255, 182, 193) ' 연한 빨강
                ws.Cells(targetRow, statusCol).Font.Color = RGB(150, 0, 0)
            Case "pending"
                ws.Cells(targetRow, statusCol).Interior.Color = RGB(255, 255, 224) ' 연한 노랑
                ws.Cells(targetRow, statusCol).Font.Color = RGB(150, 150, 0)
            Case Else
                ws.Cells(targetRow, statusCol).Interior.Color = RGB(211, 211, 211) ' 회색
                ws.Cells(targetRow, statusCol).Font.Color = RGB(100, 100, 100)
        End Select
        
        ' 셀 서식 설정
        ws.Cells(targetRow, statusCol).HorizontalAlignment = xlCenter
        ws.Cells(targetRow, statusCol).Font.Bold = True
        ws.Cells(targetRow, statusCol).Borders.LineStyle = xlContinuous
        
        ' 업데이트 시간도 기록 (다음 열에)
        If statusCol + 1 <= 256 Then ' 엑셀 열 제한 확인
            ws.Cells(1, statusCol + 1).Value = "업데이트시간"
            ws.Cells(1, statusCol + 1).Font.Bold = True
            ws.Cells(targetRow, statusCol + 1).Value = Format(Now(), "mm/dd hh:mm")
            ws.Cells(targetRow, statusCol + 1).Font.Size = 8
            ws.Cells(targetRow, statusCol + 1).HorizontalAlignment = xlCenter
        End If
        
        ' 성공 로그를 결과 영역에 표시
        If ws.Name = "재무리포트_대시보드" Then
            ws.Range("E9").Value = "✅ 승인상태 업데이트 성공!" & vbCrLf & _
                                   "월: " & monthText & vbCrLf & _
                                   "행: " & targetRow & vbCrLf & _
                                   "열: " & statusCol & vbCrLf & _
                                   "상태: " & statusText & vbCrLf & _
                                   "시간: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        End If
    Else
        ' 해당 월을 찾지 못한 경우
        MsgBox "⚠️ " & monthText & " 데이터를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "테이블에 해당 월 데이터가 있는지 확인하세요.", vbExclamation, "월 데이터 없음"
        
        ' 실패 로그를 결과 영역에 표시
        If ws.Name = "재무리포트_대시보드" Then
            ws.Range("E9").Value = "❌ 승인상태 업데이트 실패!" & vbCrLf & _
                                   "월: " & monthText & vbCrLf & _
                                   "원인: 해당 월 데이터를 찾을 수 없음" & vbCrLf & _
                                   "시간: " & Format(Now(), "yyyy-mm-dd hh:mm:ss")
        End If
    End If
End Sub

' 월별 테이블에서 승인상태 열 찾기 또는 생성
Function 월별테이블_승인상태열찾기(ws As Worksheet, lastCol As Long) As Long
    Dim i As Long
    Dim foundCol As Long
    Dim headerRow As Long
    
    ' 헤더 행 찾기 (보통 1행 또는 승인상태라는 텍스트가 있는 행)
    headerRow = 1
    
    ' 기존 "승인상태" 열 찾기
    For i = 1 To lastCol
        If InStr(ws.Cells(headerRow, i).Value, "승인상태") > 0 Or _
           InStr(ws.Cells(headerRow, i).Value, "승인") > 0 Then
            foundCol = i
            Exit For
        End If
    Next i
    
    ' 승인상태 열이 없으면 새로 생성
    If foundCol = 0 Then
        foundCol = lastCol + 1
        ws.Cells(headerRow, foundCol).Value = "승인상태"
        ws.Cells(headerRow, foundCol).Font.Bold = True
        ws.Cells(headerRow, foundCol).HorizontalAlignment = xlCenter
        ws.Cells(headerRow, foundCol).Interior.Color = RGB(200, 200, 255) ' 연한 파랑
        ws.Cells(headerRow, foundCol).Borders.LineStyle = xlContinuous
    End If
    
    월별테이블_승인상태열찾기 = foundCol
End Function

' ===== 데이터 전송 관련 함수 =====

' 재무 데이터를 API로 전송
Function SendFinanceDataToAPI(year As Integer, month As Integer) As Boolean
    Dim http As Object
    Dim url As String
    Dim jsonData As String
    Dim response As String
    Dim confirmMsg As String
    
    On Error GoTo ErrorHandler
    
    ' WinHTTP 객체 생성
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' API URL 설정
    url = API_BASE_URL & "/reports/submit"
    
    ' 엑셀에서 재무 데이터 읽기
    Dim financeData As FinanceData
    financeData = ReadFinanceDataFromCells()
    
    ' JSON 데이터 생성
    jsonData = "{"
    jsonData = jsonData & """year"": " & year & ","
    jsonData = jsonData & """month"": " & month & ","
    jsonData = jsonData & """salesRevenue"": " & financeData.salesRevenue & ","
    jsonData = jsonData & """otherIncome"": " & financeData.otherIncome & ","
    jsonData = jsonData & """rentExpense"": " & financeData.rentExpense & ","
    jsonData = jsonData & """laborExpense"": " & financeData.laborExpense & ","
    jsonData = jsonData & """materialExpense"": " & financeData.materialExpense & ","
    jsonData = jsonData & """operatingExpense"": " & financeData.operatingExpense & ","
    jsonData = jsonData & """otherExpense"": " & financeData.otherExpense & ","
    jsonData = jsonData & """cashBalance"": " & financeData.cashBalance & ","
    jsonData = jsonData & """submittedBy"": """ & Application.UserName & """"
    jsonData = jsonData & "}"
    
    ' 전송 전 파라미터 확인 Alert
    confirmMsg = "📤 데이터 전송 파라미터:" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "🌐 URL: " & url & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "📋 전송 데이터:" & vbCrLf
    confirmMsg = confirmMsg & "년도: " & year & vbCrLf
    confirmMsg = confirmMsg & "월: " & month & vbCrLf
    confirmMsg = confirmMsg & "매출: " & Format(financeData.salesRevenue, "#,##0") & "원" & vbCrLf
    confirmMsg = confirmMsg & "기타수입: " & Format(financeData.otherIncome, "#,##0") & "원" & vbCrLf
    confirmMsg = confirmMsg & "임대료: " & Format(financeData.rentExpense, "#,##0") & "원" & vbCrLf
    confirmMsg = confirmMsg & "인건비: " & Format(financeData.laborExpense, "#,##0") & "원" & vbCrLf
    confirmMsg = confirmMsg & "재료비: " & Format(financeData.materialExpense, "#,##0") & "원" & vbCrLf
    confirmMsg = confirmMsg & "운영비: " & Format(financeData.operatingExpense, "#,##0") & "원" & vbCrLf
    confirmMsg = confirmMsg & "기타비용: " & Format(financeData.otherExpense, "#,##0") & "원" & vbCrLf
    confirmMsg = confirmMsg & "현금잔고: " & Format(financeData.cashBalance, "#,##0") & "원" & vbCrLf
    confirmMsg = confirmMsg & "전송자: " & Application.UserName & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "📜 JSON 파라미터:" & vbCrLf
    confirmMsg = confirmMsg & jsonData & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "이 데이터를 서버로 전송하시겠습니까?"
    
    ' 전송 확인 Dialog
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "🚀 데이터 전송 확인") = vbNo Then
        SendFinanceDataToAPI = False
        Exit Function
    End If
    
    ' HTTP 요청 설정 및 전송
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    
    ' 요청 전송
    http.Send jsonData
    
    ' 응답 받기
    response = http.ResponseText
    
    ' 전송 후 응답 Alert
    Dim responseMsg As String
    responseMsg = "📡 서버 응답:" & vbCrLf & vbCrLf
    responseMsg = responseMsg & "🌐 HTTP 상태코드: " & http.Status & vbCrLf
    responseMsg = responseMsg & "📝 응답 헤더:" & vbCrLf
    responseMsg = responseMsg & "Content-Type: " & http.GetResponseHeader("Content-Type") & vbCrLf & vbCrLf
    responseMsg = responseMsg & "📋 응답 내용 (JSON):" & vbCrLf
    responseMsg = responseMsg & response & vbCrLf & vbCrLf
    
    ' 응답 확인
    If http.Status = 200 Then
        responseMsg = responseMsg & "✅ 전송 결과: 성공!"
        ' JSON 응답에서 success 필드 확인
        If InStr(response, """success"":true") > 0 Then
            SendFinanceDataToAPI = True
        Else
            SendFinanceDataToAPI = False
            responseMsg = responseMsg & vbCrLf & "⚠️ 주의: 서버에서 처리 오류 발생"
        End If
    Else
        SendFinanceDataToAPI = False
        responseMsg = responseMsg & "❌ 전송 결과: 실패!"
        responseMsg = responseMsg & vbCrLf & "오류 상태: HTTP " & http.Status
    End If
    
    ' 응답 결과 표시
    MsgBox responseMsg, vbInformation, "📡 전송 완료"
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    SendFinanceDataToAPI = False
    Set http = Nothing
    
    ' 오류 발생 시 Alert
    MsgBox "❌ 데이터 전송 중 오류 발생!" & vbCrLf & vbCrLf & _
           "오류 내용: " & Err.Description & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & vbCrLf & _
           "네트워크 연결 및 서버 상태를 확인하세요.", vbCritical, "🚨 전송 오류"
End Function



' 엑셀 시트들에서 재무 데이터 읽기 (다른 탭들에서 자동으로 가져오기)
Function ReadFinanceDataFromCells() As FinanceData
    Dim data As FinanceData
    Dim year As Integer
    Dim month As Integer
    
    ' 전송할 년도/월 가져오기
    year = GetCurrentYear()
    month = GetCurrentMonth()
    
    ' 각 시트에서 해당 월 데이터 읽어오기 (실제 시트 데이터에 맞게)
    data.salesRevenue = GetSalesRevenueFromSheets(year, month)      ' 매출입금
    data.otherIncome = GetOtherIncomeFromSheets(year, month)        ' 기타입금
    data.rentExpense = GetExpenseFromSheets(year, month, "비용결제")  ' 비용결제에서 임대료 부분
    data.laborExpense = GetExpenseFromSheets(year, month, "비용결제") ' 비용결제에서 인건비 부분  
    data.materialExpense = GetExpenseFromSheets(year, month, "비용결제") ' 비용결제에서 재료비 부분
    data.operatingExpense = GetExpenseFromSheets(year, month, "비용결제") ' 비용결제에서 운영비 부분
    data.otherExpense = GetExpenseFromSheets(year, month, "외상대")   ' 외상대 결제
    data.cashBalance = GetCashBalanceFromSheets(year, month)        ' 현금잔고
    
    ReadFinanceDataFromCells = data
End Function

' 시트들에서 매출 데이터 가져오기 (시트 순서로 접근)
Function GetSalesRevenueFromSheets(year As Integer, month As Integer) As Double
    Dim totalSales As Double
    Dim ws As Worksheet
    
    totalSales = 0
    
    ' 1. 첫 번째 시트 (정리표)에서 매출 데이터 가져오기
    On Error Resume Next
    Set ws = Worksheets(2) ' 2번 시트 = 20~25년 정리표
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ' 해당 년도와 월을 찾아서 매출 데이터 가져오기
        totalSales = totalSales + FindMonthlyDataInSheet(ws, year, month, "매출입금", "매출")
    End If
    
    ' 3. 세 번째 시트 (통장)에서 해당 월의 매출입금 합계 가져오기
    On Error Resume Next
    Set ws = Worksheets(3) ' 3번 시트 = 통장 (순서가 밀렸음)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        totalSales = totalSales + SumMonthlyTransactions(ws, year, month, "매출입금")
    End If
    
    GetSalesRevenueFromSheets = totalSales
End Function

' 시트들에서 기타수입 데이터 가져오기 (시트 순서로 접근)
Function GetOtherIncomeFromSheets(year As Integer, month As Integer) As Double
    Dim totalIncome As Double
    Dim ws As Worksheet
    
    totalIncome = 0
    
    ' 두 번째 시트 (정리표)에서 기타수입 찾기
    On Error Resume Next
    Set ws = Worksheets(2) ' 2번 시트 = 20~25년 정리표
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        totalIncome = FindMonthlyDataInSheet(ws, year, month, "기타입금", "기타")
    End If
    
    GetOtherIncomeFromSheets = totalIncome
End Function

' 시트들에서 비용 데이터 가져오기 (비용결제 또는 외상대에서)
Function GetExpenseFromSheets(year As Integer, month As Integer, expenseType As String) As Double
    Dim totalExpense As Double
    Dim ws As Worksheet
    
    totalExpense = 0
    
    ' 두 번째 시트 (정리표)에서 해당 비용 찾기
    On Error Resume Next
    Set ws = Worksheets(2) ' 2번 시트 = 20~25년 정리표
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ' 비용결제 또는 외상대에서 데이터 찾기
        totalExpense = FindMonthlyDataInSheet(ws, year, month, expenseType, expenseType)
    End If
    
    GetExpenseFromSheets = totalExpense
End Function

' 시트들에서 임대료 데이터 가져오기 (호환성을 위해 유지)
Function GetRentExpenseFromSheets(year As Integer, month As Integer) As Double
    GetRentExpenseFromSheets = GetExpenseFromSheets(year, month, "비용결제")
End Function

' 시트들에서 인건비 데이터 가져오기 (호환성을 위해 유지)
Function GetLaborExpenseFromSheets(year As Integer, month As Integer) As Double
    GetLaborExpenseFromSheets = GetExpenseFromSheets(year, month, "비용결제")
End Function

' 시트들에서 재료비 데이터 가져오기 (호환성을 위해 유지)
Function GetMaterialExpenseFromSheets(year As Integer, month As Integer) As Double
    GetMaterialExpenseFromSheets = GetExpenseFromSheets(year, month, "비용결제")
End Function

' 시트들에서 운영비 데이터 가져오기 (호환성을 위해 유지)
Function GetOperatingExpenseFromSheets(year As Integer, month As Integer) As Double
    GetOperatingExpenseFromSheets = GetExpenseFromSheets(year, month, "비용결제")
End Function

' 시트들에서 기타비용 데이터 가져오기 (호환성을 위해 유지)
Function GetOtherExpenseFromSheets(year As Integer, month As Integer) As Double
    GetOtherExpenseFromSheets = GetExpenseFromSheets(year, month, "외상대")
End Function

' 시트들에서 현금잔고 데이터 가져오기
Function GetCashBalanceFromSheets(year As Integer, month As Integer) As Double
    Dim cashBalance As Double
    Dim ws As Worksheet
    
    cashBalance = 0
    
    ' 1. 네 번째 시트 (캐시플로우)에서 찾기 (순서가 밀림)
    On Error Resume Next
    Set ws = Worksheets(4) ' 4번 시트 = 캐시플로우 (순서가 밀렸음)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        cashBalance = FindMonthlyDataInSheet(ws, year, month, "현금잔고", "잔고")
    End If
    
    ' 2. 캐시플로우에서 못 찾으면 세 번째 시트 (통장)에서 마지막 잔액 찾기
    If cashBalance = 0 Then
        On Error Resume Next
        Set ws = Worksheets(3) ' 3번 시트 = 통장 (순서가 밀렸음)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            cashBalance = GetLastCashBalanceFromBankSheet(ws, year, month)
        End If
    End If
    
    GetCashBalanceFromSheets = cashBalance
End Function

' 시트에서 해당 년월의 특정 항목 데이터 찾기 (실제 시트 구조 기반)
Function FindMonthlyDataInSheet(ws As Worksheet, targetYear As Integer, targetMonth As Integer, searchTerm1 As String, searchTerm2 As String) As Double
    Dim result As Double
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Variant
    Dim yearRow As Long
    Dim dataRow As Long
    Dim monthCol As Long
    
    On Error GoTo ErrorHandler
    
    result = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 안전한 범위 제한
    If lastRow > 1000 Then lastRow = 1000
    
    ' 1단계: 해당 년도 행 찾기 (A열에서 "2025년" 검색)
    For i = 1 To lastRow
        On Error Resume Next
        cellValue = ws.Cells(i, 1).Value
        On Error GoTo ErrorHandler
        
        If CStr(cellValue) = CStr(targetYear) & "년" Then
            yearRow = i
            Exit For
        End If
    Next i
    
    If yearRow = 0 Then GoTo ErrorHandler
    
    ' 2단계: 해당 월 열 찾기 (1월=B열, 2월=C열, ..., 12월=M열)
    monthCol = targetMonth + 1 ' 1월=B열(2), 2월=C열(3), ..., 6월=G열(7)
    
    ' 3단계: 해당 항목 행 찾기 (년도 행 다음부터 5행 정도 범위)
    For i = yearRow + 1 To yearRow + 5
        On Error Resume Next
        cellValue = ws.Cells(i, 1).Value
        On Error GoTo ErrorHandler
        
        If InStr(CStr(cellValue), searchTerm1) > 0 Or _
           InStr(CStr(cellValue), searchTerm2) > 0 Then
            dataRow = i
            Exit For
        End If
    Next i
    
    ' 4단계: 데이터 가져오기
    If dataRow > 0 And monthCol > 0 Then
        On Error Resume Next
        cellValue = ws.Cells(dataRow, monthCol).Value
        On Error GoTo ErrorHandler
        
        If IsNumeric(cellValue) Then
            result = CDbl(cellValue)
        End If
    End If
    
    FindMonthlyDataInSheet = result
    Exit Function
    
ErrorHandler:
    FindMonthlyDataInSheet = 0
End Function

' 통장 시트에서 해당 월의 거래 합계 구하기
Function SumMonthlyTransactions(ws As Worksheet, targetYear As Integer, targetMonth As Integer, transactionType As String) As Double
    Dim result As Double
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Variant
    Dim dateValue As Variant
    Dim amountValue As Variant
    Dim transactionValue As Variant
    
    On Error GoTo ErrorHandler
    
    result = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 안전한 범위 제한 (최대 10000행까지만)
    If lastRow > 10000 Then lastRow = 10000
    
    ' 날짜 열과 금액 열, 거래 유형 열 찾기
    For i = 2 To lastRow ' 헤더 제외
        ' 날짜 확인 (A열에 날짜가 있다고 가정)
        On Error Resume Next
        dateValue = ws.Cells(i, 1).Value
        On Error GoTo ErrorHandler
        
        If IsDate(dateValue) Then
            If Year(CDate(dateValue)) = targetYear And Month(CDate(dateValue)) = targetMonth Then
                ' 거래 유형 확인 (C열 또는 D열에 거래 유형이 있다고 가정)
                On Error Resume Next
                transactionValue = ws.Cells(i, 3).Value & " " & ws.Cells(i, 4).Value
                On Error GoTo ErrorHandler
                
                If InStr(CStr(transactionValue), transactionType) > 0 Then
                    ' 금액 더하기 (B열에 금액이 있다고 가정)
                    On Error Resume Next
                    amountValue = ws.Cells(i, 2).Value
                    On Error GoTo ErrorHandler
                    
                    If IsNumeric(amountValue) Then
                        result = result + CDbl(amountValue)
                    End If
                End If
            End If
        End If
    Next i
    
    SumMonthlyTransactions = result
    Exit Function
    
ErrorHandler:
    SumMonthlyTransactions = 0
End Function

' 통장 시트에서 해당 월 마지막 현금잔고 가져오기
Function GetLastCashBalanceFromBankSheet(ws As Worksheet, targetYear As Integer, targetMonth As Integer) As Double
    Dim result As Double
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim dateValue As Variant
    Dim balanceValue As Variant
    
    On Error GoTo ErrorHandler
    
    result = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' 안전한 범위 제한
    If lastRow > 10000 Then lastRow = 10000
    If lastCol > 50 Then lastCol = 50
    
    ' 해당 월의 마지막 잔액 찾기
    For i = lastRow To 2 Step -1 ' 뒤에서부터 찾기
        On Error Resume Next
        dateValue = ws.Cells(i, 1).Value
        On Error GoTo ErrorHandler
        
        If IsDate(dateValue) Then
            If Year(CDate(dateValue)) = targetYear And Month(CDate(dateValue)) = targetMonth Then
                ' 잔액 열 찾기 (E열부터 마지막 열까지 검색)
                For j = 5 To lastCol ' E열(5)부터 시작
                    On Error Resume Next
                    balanceValue = ws.Cells(i, j).Value
                    On Error GoTo ErrorHandler
                    
                    If IsNumeric(balanceValue) And CDbl(balanceValue) > 0 Then
                        result = CDbl(balanceValue)
                        GetLastCashBalanceFromBankSheet = result
                        Exit Function
                    End If
                Next j
            End If
        End If
    Next i
    
    GetLastCashBalanceFromBankSheet = result
    Exit Function
    
ErrorHandler:
    GetLastCashBalanceFromBankSheet = 0
End Function

' 셀 값을 안전하게 가져오기 (숫자가 아니면 기본값 반환)
Function GetCellValue(cellAddress As String, defaultValue As Variant) As Variant
    Dim cellValue As Variant
    cellValue = Range(cellAddress).Value
    
    If IsNumeric(cellValue) Then
        GetCellValue = CDbl(cellValue)
    Else
        GetCellValue = defaultValue
    End If
End Function

' 특정 시트에서 셀 값을 안전하게 가져오기
Function GetCellValueFromSheet(ws As Worksheet, cellAddress As String, defaultValue As Variant) As Variant
    Dim cellValue As Variant
    cellValue = ws.Range(cellAddress).Value
    
    If IsNumeric(cellValue) Then
        GetCellValueFromSheet = CDbl(cellValue)
    Else
        GetCellValueFromSheet = defaultValue
    End If
End Function

' 재무 데이터 유효성 검사 (실제 시트들에서 데이터 확인)
Function ValidateFinanceData() As Boolean
    Dim year As Integer
    Dim month As Integer
    Dim salesRevenue As Double
    Dim hasDataSheet As Boolean
    
    ' 전송할 년도/월 가져오기
    year = GetCurrentYear()
    month = GetCurrentMonth()
    
    ' 시트 존재 여부 확인 (시트 순서로 접근)
    hasDataSheet = False
    
    ' 2번 시트 (정리표) 확인 - 20~25년 정리표
    On Error Resume Next
    If Not Worksheets(2) Is Nothing Then
        hasDataSheet = True
    End If
    On Error GoTo 0
    
    ' 3번 시트 (통장) 확인
    On Error Resume Next
    If Not Worksheets(3) Is Nothing Then
        hasDataSheet = True
    End If
    On Error GoTo 0
    
    ' 기본 유효성 검사
    If hasDataSheet Then
        ' 매출 데이터 확인
        salesRevenue = GetSalesRevenueFromSheets(year, month)
        If salesRevenue >= 0 Then
            ValidateFinanceData = True
        Else
            ValidateFinanceData = False
        End If
    Else
        ' 데이터 시트가 없으면 false
        ValidateFinanceData = False
    End If
End Function

' ===== 승인/반려 관련 함수 =====

' API로 승인/반려 정보 전송
Function SendApprovalToAPI(month As Integer, year As Integer, approvalStatus As String, memo As String) As Boolean
    Dim http As Object
    Dim url As String
    Dim jsonData As String
    Dim response As String
    Dim confirmMsg As String
    Dim responseMsg As String
    
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
    
    ' 전송 전 파라미터 확인 Alert
    confirmMsg = "📋 " & IIf(approvalStatus = "approved", "승인", "반려") & " 처리 파라미터:" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "🌐 URL: " & url & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "📊 처리 정보:" & vbCrLf
    confirmMsg = confirmMsg & "년도: " & year & vbCrLf
    confirmMsg = confirmMsg & "월: " & month & vbCrLf
    confirmMsg = confirmMsg & "상태: " & IIf(approvalStatus = "approved", "승인", "반려") & vbCrLf
    confirmMsg = confirmMsg & "메모: " & memo & vbCrLf
    confirmMsg = confirmMsg & "처리자: " & Application.UserName & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "📜 JSON 파라미터:" & vbCrLf
    confirmMsg = confirmMsg & jsonData & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "이 " & IIf(approvalStatus = "approved", "승인", "반려") & " 처리를 서버로 전송하시겠습니까?"
    
    ' 전송 확인 Dialog
    If MsgBox(confirmMsg, vbQuestion + vbYesNo, "🚀 " & IIf(approvalStatus = "approved", "승인", "반려") & " 처리 확인") = vbNo Then
        SendApprovalToAPI = False
        Exit Function
    End If
    
    ' HTTP 요청 설정 및 전송
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    
    ' 요청 전송
    http.Send jsonData
    
    ' 응답 받기
    response = http.ResponseText
    
    ' 전송 후 응답 Alert
    responseMsg = "📡 서버 응답:" & vbCrLf & vbCrLf
    responseMsg = responseMsg & "🌐 HTTP 상태코드: " & http.Status & vbCrLf
    responseMsg = responseMsg & "📝 응답 헤더:" & vbCrLf
    responseMsg = responseMsg & "Content-Type: " & http.GetResponseHeader("Content-Type") & vbCrLf & vbCrLf
    responseMsg = responseMsg & "📋 응답 내용 (JSON):" & vbCrLf
    responseMsg = responseMsg & response & vbCrLf & vbCrLf
    
    ' 응답 확인
    If http.Status = 200 Then
        responseMsg = responseMsg & "✅ " & IIf(approvalStatus = "approved", "승인", "반려") & " 처리 결과: 성공!"
        ' JSON 응답에서 success 필드 확인
        If InStr(response, """success"":true") > 0 Then
            SendApprovalToAPI = True
        Else
            SendApprovalToAPI = False
            responseMsg = responseMsg & vbCrLf & "⚠️ 주의: 서버에서 처리 오류 발생"
        End If
    Else
        SendApprovalToAPI = False
        responseMsg = responseMsg & "❌ " & IIf(approvalStatus = "approved", "승인", "반려") & " 처리 결과: 실패!"
        responseMsg = responseMsg & vbCrLf & "오류 상태: HTTP " & http.Status
    End If
    
    ' 응답 결과 표시
    MsgBox responseMsg, vbInformation, "📡 " & IIf(approvalStatus = "approved", "승인", "반려") & " 처리 완료"
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    SendApprovalToAPI = False
    Set http = Nothing
    
    ' 오류 발생 시 Alert
    MsgBox "❌ " & IIf(approvalStatus = "approved", "승인", "반려") & " 처리 중 오류 발생!" & vbCrLf & vbCrLf & _
           "오류 내용: " & Err.Description & vbCrLf & _
           "오류 번호: " & Err.Number & vbCrLf & vbCrLf & _
           "네트워크 연결 및 서버 상태를 확인하세요.", vbCritical, "🚨 처리 오류"
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

' 현재 월 가져오기 (데이터 전송용 - C4 셀)
Function GetCurrentMonth() As Integer
    Dim cellValue As Variant
    Dim ws As Worksheet
    
    ' 대시보드 시트 찾기
    On Error Resume Next
    Set ws = Worksheets("재무리포트_대시보드")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' 대시보드 시트가 없으면 현재 시트에서 C4 확인
        cellValue = Range("C4").Value
    Else
        ' 대시보드 시트의 C4에서 월 가져오기 (데이터 전송용)
        cellValue = ws.Range("C4").Value
    End If
    
    If IsNumeric(cellValue) And cellValue >= 1 And cellValue <= 12 Then
        GetCurrentMonth = CInt(cellValue)
    Else
        GetCurrentMonth = Month(Date)
    End If
End Function

' 현재 년도 가져오기 (데이터 전송용 - C3 셀)
Function GetCurrentYear() As Integer
    Dim cellValue As Variant
    Dim ws As Worksheet
    
    ' 대시보드 시트 찾기
    On Error Resume Next
    Set ws = Worksheets("재무리포트_대시보드")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' 대시보드 시트가 없으면 현재 시트에서 C3 확인
        cellValue = Range("C3").Value
    Else
        ' 대시보드 시트의 C3에서 년도 가져오기 (데이터 전송용)
        cellValue = ws.Range("C3").Value
    End If
    
    If IsNumeric(cellValue) And cellValue >= 2020 And cellValue <= 2030 Then
        GetCurrentYear = CInt(cellValue)
    Else
        GetCurrentYear = Year(Date)
    End If
End Function

' 승인상태 확인용 년도 가져오기 (B7 셀)
Function GetApprovalStatusYear() As Integer
    Dim cellValue As Variant
    Dim ws As Worksheet
    
    ' 대시보드 시트 찾기
    On Error Resume Next
    Set ws = Worksheets("재무리포트_대시보드")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' 대시보드 시트가 없으면 현재 시트에서 B7 확인
        cellValue = Range("B7").Value
    Else
        ' 대시보드 시트의 B7에서 년도 가져오기 (승인상태 확인용)
        cellValue = ws.Range("B7").Value
    End If
    
    If IsNumeric(cellValue) And cellValue >= 2020 And cellValue <= 2030 Then
        GetApprovalStatusYear = CInt(cellValue)
    Else
        GetApprovalStatusYear = Year(Date)
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

' 승인 상태 새로고침 및 셀 업데이트 (대시보드 시트에서)
Sub RefreshApprovalStatus()
    Dim month As Integer
    Dim year As Integer
    Dim status As String
    Dim statusText As String
    Dim ws As Worksheet
    
    ' 대시보드 시트 찾기
    On Error Resume Next
    Set ws = Worksheets("재무리포트_대시보드")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ActiveSheet
    End If
    
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
    
    ' 상태를 대시보드 시트의 F4 셀에 표시
    ws.Range("F4").Value = statusText
    
    ' 상태에 따라 셀 색상 변경
    Select Case status
        Case "approved"
            ws.Range("F4").Interior.Color = RGB(144, 238, 144) ' 연한 녹색
        Case "rejected"
            ws.Range("F4").Interior.Color = RGB(255, 182, 193) ' 연한 빨강
        Case "pending"
            ws.Range("F4").Interior.Color = RGB(255, 255, 224) ' 연한 노랑
        Case Else
            ws.Range("F4").Interior.Color = RGB(211, 211, 211) ' 회색
    End Select
    
    ' 서버 상태도 업데이트
    If status <> "error" Then
        ws.Range("F6").Value = "연결됨"
        ws.Range("F6").Interior.Color = RGB(144, 238, 144) ' 연한 녹색
    Else
        ws.Range("F6").Value = "연결실패"
        ws.Range("F6").Interior.Color = RGB(255, 182, 193) ' 연한 빨강
    End If
    
    ' 마지막 업데이트 시간 표시
    ws.Range("F5").Value = Format(Now(), "hh:mm:ss")
    ws.Range("F5").Interior.Color = RGB(248, 248, 248)
End Sub

' ===== 설정 및 초기화 =====

' 빠른 설정 실행 (보안 안내 포함)
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
    
    ' 5. API 연결 테스트
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

' 새로운 재무 대시보드 시트 생성 및 버튼 자동 생성 (새로운 레이아웃)
Sub 버튼_자동생성()
    Dim ws As Worksheet
    Dim wsName As String
    Dim btnDataSend As Button
    Dim btnRefresh As Button
    
    ' 새로운 시트 이름 설정
    wsName = "재무리포트_대시보드"
    
    ' 기존 시트가 있으면 삭제
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' 새로운 시트 생성
    Set ws = Worksheets.Add
    ws.Name = wsName
    ws.Activate
    
    ' 새로운 레이아웃 설정
    Call 새로운레이아웃_설정(ws)
    
    ' 기존 버튼들 삭제
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    
    ' ===== 메인 버튼들 =====
    
    ' 데이터 전송 버튼 (D3 위치)
    Set btnDataSend = ws.Buttons.Add(ws.Range("D3").Left, ws.Range("D3").Top, 80, 25)
    btnDataSend.OnAction = "데이터전송"
    btnDataSend.Caption = "데이터전송"
    btnDataSend.Font.Size = 10
    btnDataSend.Font.Bold = True
    
    ' 새로고침 버튼 (D6 위치) - 전체월 승인상태 확인 함수에 매핑
    Set btnRefresh = ws.Buttons.Add(ws.Range("D6").Left, ws.Range("D6").Top, 80, 25)
    btnRefresh.OnAction = "전체월_승인상태확인"
    btnRefresh.Caption = "새로고침"
    btnRefresh.Font.Size = 10
    btnRefresh.Font.Bold = True
    
    ' 디버깅 버튼들 추가
    Dim btnDebug As Button
    Dim btnStructure As Button
    
    ' 데이터 수집 디버깅 버튼 (F12)
    Set btnDebug = ws.Buttons.Add(ws.Range("F12").Left, ws.Range("F12").Top, 90, 25)
    btnDebug.OnAction = "데이터수집_디버깅"
    btnDebug.Caption = "🔍 데이터 디버깅"
    btnDebug.Font.Size = 9
    
    ' 시트 구조 분석 버튼 (F13)
    Set btnStructure = ws.Buttons.Add(ws.Range("F13").Left, ws.Range("F13").Top, 90, 25)
    btnStructure.OnAction = "시트구조_분석"
    btnStructure.Caption = "📋 시트 구조 분석"
    btnStructure.Font.Size = 9
    
        MsgBox "재무 리포트 대시보드가 생성되었습니다!" & vbCrLf & vbCrLf & _
           "📋 사용법:" & vbCrLf & _
           "1. C3, C4에 연도/월 입력 (데이터 전송용)" & vbCrLf & _
           "2. B7에 연도 입력 (승인상태 확인용)" & vbCrLf & _
           "3. '데이터전송' 버튼: 다른 시트에서 데이터 자동 수집하여 전송" & vbCrLf & _
           "4. '새로고침' 버튼: 전체월 승인상태 확인" & vbCrLf & vbCrLf & _
           "💡 필요한 시트: 20-25년 정리표, 통장, 캐시플로우", vbInformation, "대시보드 생성 완료"
End Sub

' 새로운 레이아웃 설정 (이미지와 동일한 구조)
Sub 새로운레이아웃_설정(ws As Worksheet)
    ' 시트 보호 해제
    ws.Unprotect
    
    ' ===== 데이터 전송 영역 =====
    
    ' A2: "대시보드에 데이터 전송" (병합)
    ws.Range("A2:D2").Merge
    ws.Range("A2").Value = "대시보드에 데이터 전송"
    ws.Range("A2").Font.Size = 12
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Interior.Color = RGB(255, 255, 0) ' 노란색
    ws.Range("A2").HorizontalAlignment = xlCenter
    ws.Range("A2").Borders.LineStyle = xlContinuous
    
    ' B3: "연도", C3: 연도 입력
    ws.Range("B3").Value = "연도"
    ws.Range("B3").Font.Bold = True
    ws.Range("C3").Value = Year(Date)
    ws.Range("C3").NumberFormat = "0"
    ws.Range("C3").HorizontalAlignment = xlCenter
    ws.Range("C3").Interior.Color = RGB(255, 255, 224)
    ws.Range("C3").Borders.LineStyle = xlContinuous
    
    ' B4: "월", C4: 월 입력
    ws.Range("B4").Value = "월"
    ws.Range("B4").Font.Bold = True
    ws.Range("C4").Value = Month(Date)
    ws.Range("C4").NumberFormat = "0"
    ws.Range("C4").HorizontalAlignment = xlCenter
    ws.Range("C4").Interior.Color = RGB(255, 255, 224)
    ws.Range("C4").Borders.LineStyle = xlContinuous
    
    ' D3: "데이터전송" 버튼 자리 (함수에서 버튼 생성)
    
    ' ===== 승인상태 영역 =====
    
    ' A6: "승인상태" (병합)
    ws.Range("A6:D6").Merge
    ws.Range("A6").Value = "승인상태"
    ws.Range("A6").Font.Size = 12
    ws.Range("A6").Font.Bold = True
    ws.Range("A6").Interior.Color = RGB(255, 255, 0) ' 노란색
    ws.Range("A6").HorizontalAlignment = xlCenter
    ws.Range("A6").Borders.LineStyle = xlContinuous
    
    ' B7: "연도", C7: 승인상태 확인용 연도
    ws.Range("B7").Value = "연도"
    ws.Range("B7").Font.Bold = True
    ws.Range("B7").Interior.Color = RGB(255, 255, 0) ' 노란색
    ws.Range("B7").Borders.LineStyle = xlContinuous
    ws.Range("B7").HorizontalAlignment = xlCenter
    ws.Range("C7").Value = Year(Date)
    ws.Range("C7").NumberFormat = "0"
    ws.Range("C7").HorizontalAlignment = xlCenter
    ws.Range("C7").Interior.Color = RGB(255, 255, 224)
    ws.Range("C7").Borders.LineStyle = xlContinuous
    
    ' D6: "새로고침" 버튼 자리 (함수에서 버튼 생성)
    
    ' ===== 월별 리스트 생성 (A8~A19: 1월~12월) =====
    
    Dim i As Integer
    For i = 1 To 12
        ws.Range("A" & (7 + i)).Value = i & "월"
        ws.Range("A" & (7 + i)).Font.Bold = True
        ws.Range("A" & (7 + i)).Borders.LineStyle = xlContinuous
        ws.Range("A" & (7 + i)).HorizontalAlignment = xlCenter
        ws.Range("A" & (7 + i)).Interior.Color = RGB(248, 248, 248)
        
        ' B열: 승인상태가 들어갈 자리
        ws.Range("B" & (7 + i)).Value = ""
        ws.Range("B" & (7 + i)).Borders.LineStyle = xlContinuous
        ws.Range("B" & (7 + i)).HorizontalAlignment = xlCenter
        ws.Range("B" & (7 + i)).Interior.Color = RGB(255, 255, 255)
    Next i
    
    ' ===== 열 너비 조정 =====
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    ws.Columns("F").ColumnWidth = 12
    
    ' 데이터 유효성 검사 추가
    With ws.Range("C3").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="2020", Formula2:="2030"
        .ErrorTitle = "년도 입력 오류"
        .ErrorMessage = "2020년부터 2030년 사이의 값을 입력하세요."
    End With
    
    With ws.Range("C4").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="1", Formula2:="12"
        .ErrorTitle = "월 입력 오류"
        .ErrorMessage = "1월부터 12월 사이의 값을 입력하세요."
    End With
    
    With ws.Range("B7").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="2020", Formula2:="2030"
        .ErrorTitle = "년도 입력 오류"
                 .ErrorMessage = "2020년부터 2030년 사이의 값을 입력하세요."
     End With
     
     ' ===== 안내 메시지 영역 =====
     
     ' F2: "자동 데이터 수집"
     ws.Range("F2").Value = "📊 자동 데이터 수집"
     ws.Range("F2").Font.Size = 11
     ws.Range("F2").Font.Bold = True
     ws.Range("F2").Interior.Color = RGB(200, 255, 200) ' 연한 녹색
     
     ' F3~F10: 데이터 소스 안내
     ws.Range("F3").Value = "데이터 소스:"
     ws.Range("F4").Value = "• 20-25년 정리표"
     ws.Range("F5").Value = "• 통장 시트"
     ws.Range("F6").Value = "• 캐시플로우 시트"
     ws.Range("F7").Value = ""
     ws.Range("F8").Value = "전송 시 해당 월의"
     ws.Range("F9").Value = "데이터를 자동으로"
     ws.Range("F10").Value = "수집하여 전송합니다."
     
     ' 안내 메시지 서식
     ws.Range("F3").Font.Bold = True
     ws.Range("F4:F6").Font.Size = 9
     ws.Range("F4:F6").Interior.Color = RGB(245, 245, 245) ' 연한 회색
     ws.Range("F8:F10").Font.Size = 9
     ws.Range("F8:F10").Font.Color = RGB(100, 100, 100) ' 회색 글자
     ws.Range("F8:F10").Font.Italic = True
End Sub

' 대시보드 시트 기본 설정
Sub 대시보드시트_기본설정(ws As Worksheet)
    ' 시트 보호 해제
    ws.Unprotect
    
    ' 시트 제목
    ws.Range("A1").Value = "🏢 재무 리포트 대시보드"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(0, 102, 204)
    ws.Range("A1:G1").Merge
    ws.Range("A1").HorizontalAlignment = xlCenter
    
    ' 열 너비 조정
    ws.Columns("A").ColumnWidth = 15
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 15
    ws.Columns("G").ColumnWidth = 20
    
    ' 격자 표시
    ws.Cells.Borders.LineStyle = xlNone
End Sub

' 년도/월 입력 영역 설정
Sub 년도월_입력영역_설정(ws As Worksheet)
    ' 년도/월 입력 섹션
    ws.Range("A3").Value = "📅 년도/월 설정"
    ws.Range("A3").Font.Size = 12
    ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Color = RGB(204, 102, 0)
    
    ' 년도 입력
    ws.Range("A4").Value = "년도:"
    ws.Range("A4").Font.Bold = True
    ws.Range("B4").Value = Year(Date)
    ws.Range("B4").NumberFormat = "0"
    ws.Range("B4").HorizontalAlignment = xlCenter
    ws.Range("B4").Interior.Color = RGB(255, 255, 224)
    ws.Range("B4").Borders.LineStyle = xlContinuous
    
    ' 월 입력
    ws.Range("A5").Value = "월:"
    ws.Range("A5").Font.Bold = True
    ws.Range("B5").Value = Month(Date)
    ws.Range("B5").NumberFormat = "0"
    ws.Range("B5").HorizontalAlignment = xlCenter
    ws.Range("B5").Interior.Color = RGB(255, 255, 224)
    ws.Range("B5").Borders.LineStyle = xlContinuous
    
    ' 데이터 유효성 검사 추가
    With ws.Range("B4").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="2020", Formula2:="2030"
        .ErrorTitle = "년도 입력 오류"
        .ErrorMessage = "2020년부터 2030년 사이의 값을 입력하세요."
    End With
    
    With ws.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="1", Formula2:="12"
        .ErrorTitle = "월 입력 오류"
        .ErrorMessage = "1월부터 12월 사이의 값을 입력하세요."
    End With
End Sub

' 재무 데이터 입력 영역 설정
Sub 재무데이터_입력영역_설정(ws As Worksheet)
    ' 재무 데이터 입력 섹션
    ws.Range("A7").Value = "💰 재무 데이터 입력"
    ws.Range("A7").Font.Size = 12
    ws.Range("A7").Font.Bold = True
    ws.Range("A7").Font.Color = RGB(0, 153, 0)
    
    ' 매출 섹션
    ws.Range("A8").Value = "📈 매출"
    ws.Range("A8").Font.Bold = True
    ws.Range("A8").Font.Color = RGB(0, 102, 204)
    
    ws.Range("B9").Value = "매출:"
    ws.Range("B10").Value = "기타수입:"
    
    ' 지출 섹션
    ws.Range("A11").Value = "📉 지출"
    ws.Range("A11").Font.Bold = True
    ws.Range("A11").Font.Color = RGB(204, 0, 0)
    
    ws.Range("B12").Value = "임대료:"
    ws.Range("B13").Value = "인건비:"
    ws.Range("B14").Value = "재료비:"
    ws.Range("B15").Value = "운영비:"
    ws.Range("B16").Value = "기타비용:"
    
    ' 현금 섹션
    ws.Range("A17").Value = "💵 현금"
    ws.Range("A17").Font.Bold = True
    ws.Range("A17").Font.Color = RGB(153, 102, 0)
    
    ws.Range("B18").Value = "현금잔고:"
    
    ' 기본값 및 서식 설정
    Dim inputRanges As Range
    Set inputRanges = ws.Range("C9:C10,C12:C16,C18")
    
    With inputRanges
        .Value = 0
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlRight
        .Interior.Color = RGB(240, 248, 255)
        .Borders.LineStyle = xlContinuous
        .Font.Size = 10
    End With
    
    ' 라벨 서식
    ws.Range("B9:B10,B12:B16,B18").Font.Bold = True
    ws.Range("B9:B10,B12:B16,B18").HorizontalAlignment = xlRight
End Sub

' 상태 표시 영역 설정
Sub 상태표시_영역_설정(ws As Worksheet)
    ' 상태 표시 섹션
    ws.Range("E3").Value = "📊 상태 정보"
    ws.Range("E3").Font.Size = 12
    ws.Range("E3").Font.Bold = True
    ws.Range("E3").Font.Color = RGB(102, 0, 204)
    
    ' 승인 상태
    ws.Range("E4").Value = "승인상태:"
    ws.Range("E4").Font.Bold = True
    ws.Range("F4").Value = "확인 중..."
    ws.Range("F4").HorizontalAlignment = xlCenter
    ws.Range("F4").Interior.Color = RGB(255, 255, 224)
    ws.Range("F4").Borders.LineStyle = xlContinuous
    
    ' 마지막 전송 시간
    ws.Range("E5").Value = "전송시간:"
    ws.Range("E5").Font.Bold = True
    ws.Range("F5").Value = "-"
    ws.Range("F5").HorizontalAlignment = xlCenter
    ws.Range("F5").Interior.Color = RGB(248, 248, 248)
    ws.Range("F5").Borders.LineStyle = xlContinuous
    
    ' 서버 상태
    ws.Range("E6").Value = "서버상태:"
    ws.Range("E6").Font.Bold = True
    ws.Range("F6").Value = "미확인"
    ws.Range("F6").HorizontalAlignment = xlCenter
    ws.Range("F6").Interior.Color = RGB(248, 248, 248)
    ws.Range("F6").Borders.LineStyle = xlContinuous
    
    ' 결과 표시 영역
    ws.Range("E8").Value = "📋 처리 결과"
    ws.Range("E8").Font.Size = 12
    ws.Range("E8").Font.Bold = True
    ws.Range("E8").Font.Color = RGB(102, 0, 204)
    
    ws.Range("E9:G15").Merge
    ws.Range("E9").Value = "여기에 API 응답 결과가 표시됩니다."
    ws.Range("E9").VerticalAlignment = xlTop
    ws.Range("E9").WrapText = True
    ws.Range("E9").Interior.Color = RGB(248, 248, 248)
    ws.Range("E9").Borders.LineStyle = xlContinuous
    ws.Range("E9").Font.Size = 9
End Sub

' API 연결 확인
Sub API연결_확인()
    Dim result As String
    result = "API 서버 연결을 확인하는 중..."
    Range("D2").Value = result
    
    ' 실제 API 테스트 실행
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

' 재무 데이터 입력 템플릿 생성
Sub 재무데이터_템플릿생성()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 템플릿 레이블 생성
    ws.Range("B4").Value = "=== 재무 데이터 입력 ==="
    ws.Range("B5").Value = "매출:"
    ws.Range("B6").Value = "기타수입:"
    ws.Range("B7").Value = "--- 지출 ---"
    ws.Range("B8").Value = "임대료:"
    ws.Range("B9").Value = "인건비:"
    ws.Range("B10").Value = "재료비:"
    ws.Range("B11").Value = "운영비:"
    ws.Range("B12").Value = "기타비용:"
    ws.Range("B13").Value = "--- 현금 ---"
    ws.Range("B14").Value = "현금잔고:"
    
    ' 기본값 설정
    ws.Range("C5").Value = 0  ' 매출
    ws.Range("C6").Value = 0  ' 기타수입
    ws.Range("C8").Value = 0  ' 임대료
    ws.Range("C9").Value = 0  ' 인건비
    ws.Range("C10").Value = 0 ' 재료비
    ws.Range("C11").Value = 0 ' 운영비
    ws.Range("C12").Value = 0 ' 기타비용
    ws.Range("C14").Value = 0 ' 현금잔고
    
    ' 서식 설정
    ws.Range("B4").Font.Bold = True
    ws.Range("B5:B14").Font.Bold = True
    ws.Range("C5:C14").NumberFormat = "#,##0"
    ws.Range("C5:C14").HorizontalAlignment = xlRight
    
    ' 셀 크기 조정
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 12
End Sub

' 버튼 및 UI 설정 (한 번만 실행)
Sub 버튼설정()
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
    
    ' 라벨 설정
    Range("A1").Value = "년도:"
    Range("A2").Value = "월:"
    Range("A3").Value = "승인상태:"
    
    ' 기본값 설정
    Range("B1").Value = Year(Date)
    Range("B2").Value = Month(Date)
    
    MsgBox "버튼 설정이 완료되었습니다.", vbInformation, "설정 완료"
End Sub

' 전체 설정 실행 (이것만 실행하면 모든 설정 완료)
Sub 전체설정_실행()
    On Error Resume Next
    
    ' 1. 기본 워크시트 설정
    Range("A1").Value = "년도:"
    Range("A2").Value = "월:"
    Range("A3").Value = "승인상태:"
    Range("B1").Value = Year(Date)
    Range("B2").Value = Month(Date)
    Range("D2").Value = "확인 중..."
    
    ' 2. 재무 데이터 템플릿 생성
    Call 재무데이터_템플릿생성
    
    ' 3. 버튼 생성
    Call 버튼설정
    
    ' 4. 상태 새로고침
    Call RefreshApprovalStatus
    
    MsgBox "전체 설정이 완료되었습니다!" & vbCrLf & _
           "이제 C5~C14 셀에 재무 데이터를 입력하고" & vbCrLf & _
           "'📤 데이터전송' 버튼을 사용하세요.", vbInformation, "설정 완료"
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

' ===== 추가 테스트 함수들 =====

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
    url = "http://localhost:3001/api/test?message=VBA에서 안녕하세요!"
    
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

' 포트 연결 테스트
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
    url = "http://localhost:" & port & "/api/test"
    
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

' ===== 자동 실행 함수 =====

' 워크북 열릴 때 자동으로 상태 새로고침
Sub Auto_Open()
    RefreshApprovalStatus
End Sub

' 워크북이 활성화될 때 자동으로 상태 새로고침
Sub Workbook_Activate()
    RefreshApprovalStatus
End Sub 