' ========================================================
' 재무 데이터 전송 기능 VBA 코드
' ========================================================

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
    
    ' 확인 메시지
    If MsgBox(year & "년 " & month & "월 재무 데이터를 서버로 전송하시겠습니까?", vbQuestion + vbYesNo, "데이터 전송 확인") = vbNo Then
        Exit Sub
    End If
    
    ' 데이터 전송 실행
    result = SendFinanceDataToAPI(year, month)
    
    If result Then
        MsgBox year & "년 " & month & "월 재무 데이터가 성공적으로 전송되었습니다!" & vbCrLf & "이제 승인 대기 상태입니다.", vbInformation, "전송 완료"
        RefreshApprovalStatus
    Else
        MsgBox "데이터 전송 중 오류가 발생했습니다.", vbCritical, "전송 오류"
    End If
End Sub

' 재무 데이터를 API로 전송
Function SendFinanceDataToAPI(year As Integer, month As Integer) As Boolean
    Dim http As Object
    Dim url As String
    Dim jsonData As String
    Dim response As String
    
    On Error GoTo ErrorHandler
    
    ' WinHTTP 객체 생성
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' API URL 설정
    url = API_BASE_URL & "/reports/submit"
    
    ' 엑셀에서 재무 데이터 읽기
    Dim financeData As Object
    Set financeData = ReadFinanceDataFromCells()
    
    ' JSON 데이터 생성
    jsonData = "{"
    jsonData = jsonData & """year"": " & year & ","
    jsonData = jsonData & """month"": " & month & ","
    jsonData = jsonData & """salesRevenue"": " & financeData("salesRevenue") & ","
    jsonData = jsonData & """otherIncome"": " & financeData("otherIncome") & ","
    jsonData = jsonData & """rentExpense"": " & financeData("rentExpense") & ","
    jsonData = jsonData & """laborExpense"": " & financeData("laborExpense") & ","
    jsonData = jsonData & """materialExpense"": " & financeData("materialExpense") & ","
    jsonData = jsonData & """operatingExpense"": " & financeData("operatingExpense") & ","
    jsonData = jsonData & """otherExpense"": " & financeData("otherExpense") & ","
    jsonData = jsonData & """cashBalance"": " & financeData("cashBalance") & ","
    jsonData = jsonData & """submittedBy"": """ & Application.UserName & """"
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
            SendFinanceDataToAPI = True
        Else
            SendFinanceDataToAPI = False
        End If
    Else
        SendFinanceDataToAPI = False
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    SendFinanceDataToAPI = False
    Set http = Nothing
End Function

' 엑셀 셀에서 재무 데이터 읽기
Function ReadFinanceDataFromCells() As Object
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    
    ' 재무 데이터 셀 위치 (예시 - 실제 셀 위치에 맞게 수정하세요)
    data("salesRevenue") = GetCellValue("C5", 0)      ' 매출
    data("otherIncome") = GetCellValue("C6", 0)       ' 기타수입
    data("rentExpense") = GetCellValue("C8", 0)       ' 임대료
    data("laborExpense") = GetCellValue("C9", 0)      ' 인건비
    data("materialExpense") = GetCellValue("C10", 0)  ' 재료비
    data("operatingExpense") = GetCellValue("C11", 0) ' 운영비
    data("otherExpense") = GetCellValue("C12", 0)     ' 기타비용
    data("cashBalance") = GetCellValue("C14", 0)      ' 현금잔고
    
    Set ReadFinanceDataFromCells = data
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

' 재무 데이터 유효성 검사
Function ValidateFinanceData() As Boolean
    Dim salesRevenue As Variant
    
    ' 최소한 매출 데이터는 있어야 함
    salesRevenue = Range("C5").Value
    
    If IsNumeric(salesRevenue) And salesRevenue >= 0 Then
        ValidateFinanceData = True
    Else
        ValidateFinanceData = False
    End If
End Function

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
    
    MsgBox "재무 데이터 입력 템플릿이 생성되었습니다." & vbCrLf & "C5~C14 셀에 데이터를 입력하세요.", vbInformation, "템플릿 생성 완료"
End Sub

' 데이터 전송 전 미리보기
Sub 데이터전송_미리보기()
    Dim year As Integer
    Dim month As Integer
    Dim financeData As Object
    Dim msg As String
    
    year = GetCurrentYear()
    month = GetCurrentMonth()
    Set financeData = ReadFinanceDataFromCells()
    
    ' 미리보기 메시지 구성
    msg = year & "년 " & month & "월 재무 데이터 미리보기:" & vbCrLf & vbCrLf
    msg = msg & "매출: " & Format(financeData("salesRevenue"), "#,##0") & "원" & vbCrLf
    msg = msg & "기타수입: " & Format(financeData("otherIncome"), "#,##0") & "원" & vbCrLf
    msg = msg & "총 매출: " & Format(financeData("salesRevenue") + financeData("otherIncome"), "#,##0") & "원" & vbCrLf & vbCrLf
    msg = msg & "임대료: " & Format(financeData("rentExpense"), "#,##0") & "원" & vbCrLf
    msg = msg & "인건비: " & Format(financeData("laborExpense"), "#,##0") & "원" & vbCrLf
    msg = msg & "재료비: " & Format(financeData("materialExpense"), "#,##0") & "원" & vbCrLf
    msg = msg & "운영비: " & Format(financeData("operatingExpense"), "#,##0") & "원" & vbCrLf
    msg = msg & "기타비용: " & Format(financeData("otherExpense"), "#,##0") & "원" & vbCrLf
    msg = msg & "총 지출: " & Format(financeData("rentExpense") + financeData("laborExpense") + financeData("materialExpense") + financeData("operatingExpense") + financeData("otherExpense"), "#,##0") & "원" & vbCrLf & vbCrLf
    msg = msg & "현금잔고: " & Format(financeData("cashBalance"), "#,##0") & "원" & vbCrLf & vbCrLf
    msg = msg & "이 데이터를 전송하시겠습니까?"
    
    If MsgBox(msg, vbQuestion + vbYesNo, "데이터 전송 미리보기") = vbYes Then
        Call 데이터전송
    End If
End Sub 