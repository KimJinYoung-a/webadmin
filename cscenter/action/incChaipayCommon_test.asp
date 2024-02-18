<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'+-------------------------------------------------------------------------------------------------------------+
'|                                  카 카 오 페 이   결 제   함 수 선 언                                       |
'+----------------------------------------------+--------------------------------------------------------------+
'|                함 수 명                      |                          기    능                            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallChaiPayCancel                          | 전체 취소 요청 호출                                          |
'| (결제번호,취소금액)                           | - 반환값 : 취소내역 Object (에러시 code: ERR)               |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallChaiPayPartialCancel                    | 부분 취소 요청 호출                                          |
'| (결제번호, 결제잔액, 취소금액)                | - 반환값 : 취소내역 Object (에러시 code: ERR)               |
'+----------------------------------------------+--------------------------------------------------------------+

Dim ChaiPay_API_URL		        	    '// 백엔드 API 호출 URL
Dim ChaiPay_Check_API_URL				'// 결제내역 API 호출 URL
Dim ChaiPay_PublicKey                    '// Public-API-Key (공개키)
Dim ChaiPay_PrivateKey                    '// Private-API-Key (비밀키)

if (application("Svr_Info")="Dev") then
    ChaiPay_API_URL = "https://api-staging.chai.finance"
	ChaiPay_PublicKey = "459aae6c-2212-4e2f-9f81-d662e4df4709"
    ChaiPay_PrivateKey = "66eebc2f-5c33-4c63-8443-373b07be0c2d"
else
	ChaiPay_API_URL = "https://api.chai.finance"
	ChaiPay_PublicKey = "c8aff30b-cc9b-4d03-bb4b-168e8db10d30"
    ChaiPay_PrivateKey = "492e0396-bae7-47ff-b23c-0f3f882907ff"
end if

'=================================================
'// 전체 취소 요청 호출 함수
Function fnCallChaiPayCancel(npId, idempotencyKey, cancelAmount, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	sURL = ChaiPay_API_URL & "/v1/payment/" & npId & "/cancel"

	'// 전송 테이터 Setting
	sParam = "cancelAmount=" & cancelAmount										         '결제 고유번호. 20자.
    
	Dim oRstMsg
	Set oRstMsg = new cResultChai											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Private-API-Key", ChaiPay_PrivateKey
    oXML.setRequestHeader "Idempotency-Key", idempotencyKey
	oXML.send sParam	'파라메터 전송
	response.write oXML.responseText & "<br>"
	response.end
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제취소인증 통신 오류[004]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제취소인증 내부 오류[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallChaiPayCancel = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.status="canceled" then
		oRstMsg.code = "Success"
		Set fnCallChaiPayCancel = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제취소인증 오류. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제취소 인증결과 파징 오류[006]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallChaiPayCancel = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// 취소요청 호출 함수
Function fnCallChaiPayPartialCancel(npId, idempotencyKey, mainpaymentorg ,cancelAmount, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	sURL = ChaiPay_API_URL & "/v1/payment/" & npId & "/cancel"

	'// 전송 테이터 Setting
	sParam = "cancelAmount=" & cancelAmount					'취소금액
	sParam = sParam &"&checkoutAmount=" & mainpaymentorg  '취소요청시의 결제잔액(중복취소 방지용)

	Dim oRstMsg
	Set oRstMsg = new cResultChai											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Private-API-Key", ChaiPay_PrivateKey
    oXML.setRequestHeader "Idempotency-Key", idempotencyKey
	oXML.send sParam	'파라메터 전송
'response.write mainpaymentorg & "<br>"
response.write oXML.status & "<br>"
response.write oXML.responseText & "<br>"
'response.end
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제취소인증 통신 오류[004]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제취소인증 내부 오류[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallChaiPayPartialCancel = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.status="canceled" then
		oRstMsg.code = "Success"
		Set fnCallChaiPayPartialCancel = oResult			'파징된 정보 반환
	elseif oResult.status="confirmed" then
		oRstMsg.code = "Success"
		Set fnCallChaiPayPartialCancel = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제취소인증 오류. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제취소 인증결과 파징 오류[006]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallChaiPayPartialCancel = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'==================================================
'// 결과 반환용 객체 선언
Class cResultChai
	public code
	public message

	Private Sub Class_Initialize()
		code = ""
		message = ""
	End Sub

	Private Sub Class_Terminate()
	End Sub
end Class
%>