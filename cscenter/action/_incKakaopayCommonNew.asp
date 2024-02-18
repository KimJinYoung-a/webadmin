<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'+-------------------------------------------------------------------------------------------------------------+
'|                                  카 카 오 페 이   결 제   함 수 선 언                                       |
'+----------------------------------------------+--------------------------------------------------------------+
'|                함 수 명                      |                          기    능                            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPayCancel                        | 전체 취소 요청 호출                                          |
'| (결제번호,취소금액)                           | - 반환값 : 취소내역 Object (에러시 code: ERR)               |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPayPartialCancel                  | 부분 취소 요청 호출                                          |
'| (결제번호,취소금액)                           | - 반환값 : 취소내역 Object (에러시 code: ERR)               |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPayFileUrl  		                 | 결제 내역 파일 url 요청 호출                                  |
'| (조회날짜)              				         | - 반환값 : 결제내역 파일 경로 (에러시 code: ERR)            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPayCheckList                      | 결제 내역 요청 호출                                          |
'| (파일URL)              				        | - 반환값 : 결제내역 Json (에러시 code: ERR)                 |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPaySettlementsFileUrl              | 정산 내역 요청 호출                                          |
'| (조회날짜)              				        | - 반환값 : 정산 내역 파일 경로 (에러시 code: ERR)            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPaySettlementsCheckList           | 결제내역 요청 호출                                           |
'| (파일URL)              				        | - 반환값 : 정산 내역 Json (에러시 code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+

Dim KPay_API_URL		        	  '// 백엔드 API 호출 URL
Dim KPay_Check_API_URL				'// 결제내역 API 호출 URL
Dim KPay_PartnerID                     '// 카카오페이 파트너 ID
Dim KPay_ClientId, KPay_ClientKey      '// 텐바이텐 상점 ID 및 시크릿키
Dim KPay_FileKey					   '//결재내역, 정산 파일 접근
Dim KPay_PaymentsFileBID  			  '//결재내역 BID(bucket_id)
Dim KPay_SettlementsFileBID			  '//정산내역 BID(bucket_id)

'카카오페이 기본 경로
KPay_API_URL = "https://kapi.kakao.com"
KPay_Check_API_URL = "https://biz-api.kakaopay.com"
if (application("Svr_Info")="Dev") then
	'인증키
	KPay_PartnerID = "TC0ONETIME"							 '가맹점 코드. 10자
	KPay_ClientId = "cb8b50980734335b667cffb32781e5a1"      'Admin Key
	KPay_ClientKey = "d0e01cbc77f59fa73969f46ae83dd9ca"     'RestAPI Key
	KPay_PaymentsFileBID = "B656960246"
	KPay_SettlementsFileBID = "B997240247"
    KPay_FileKey = "cc061a13c0e1f1580c8c4168439b916d585c61ffb219cd65d8cc2cbfd7f74766" '파일 접근 키

else
	'인증키
	KPay_PartnerID = "C371930065"							 '가맹점 코드. 10자
	KPay_ClientId = "5a1f82cd75e1002b529edc6f213d875a"      'Admin Key
	KPay_ClientKey = "b4e7e01a2ade8ecedc5c6944941ffbd4"     'RestAPI Key
	KPay_PaymentsFileBID = "B656960246"
	KPay_SettlementsFileBID = "B997240247"
    KPay_FileKey = "cc061a13c0e1f1580c8c4168439b916d585c61ffb219cd65d8cc2cbfd7f74766"
end if

'=================================================
'// 취소요청 호출 함수
Function fnCallKakaoPayCancel(npId, cancelAmount, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	sURL = KPay_API_URL & "/v1/payment/cancel"

	'// 전송 테이터 Setting
	sParam = "tid=" & npId										         '결제 고유번호. 20자.
	sParam = sParam &"&cid=" & KPay_PartnerID                          '가맹점 코드. 10자
    sParam = sParam &"&cancel_amount=" & cancelAmount                '취소금액
    sParam = sParam &"&cancel_tax_free_amount=" & 0        				'취소 비과세 금액
    
	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "KakaoAK " & KPay_ClientId
	oXML.send sParam	'파라메터 전송
	'response.write oXML.responseText & "<br>"
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
		set fnCallKakaoPayCancel = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.status="CANCEL_PAYMENT" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayCancel = oResult			'파징된 정보 반환
	elseif oResult.status="PART_CANCEL_PAYMENT" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayCancel = oResult			'파징된 정보 반환
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
		set fnCallKakaoPayCancel = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// 취소요청 호출 함수
Function fnCallKakaoPayPartialCancel(npId, mainpaymentorg ,cancelAmount, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	sURL = KPay_API_URL & "/v1/payment/cancel"

	'// 전송 테이터 Setting
	sParam = "tid=" & npId										        '결제 고유번호. 20자.
	sParam = sParam &"&cid=" & KPay_PartnerID                         '가맹점 코드. 10자
    sParam = sParam &"&cancel_amount=" & cancelAmount               '취소금액
    sParam = sParam &"&cancel_tax_free_amount=" & 0        			  '취소 비과세 금액
	sParam = sParam &"&cancel_available_amount=" & mainpaymentorg   '최초 결제금액
    
	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "KakaoAK " & KPay_ClientId
	oXML.send sParam	'파라메터 전송
	'response.write oXML.responseText & "<br>"
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
		set fnCallKakaoPayPartialCancel = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.status="CANCEL_PAYMENT" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayPartialCancel = oResult			'파징된 정보 반환
	elseif oResult.status="PART_CANCEL_PAYMENT" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayPartialCancel = oResult			'파징된 정보 반환
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
		set fnCallKakaoPayPartialCancel = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// 결제내역 API 호출 함수(파일 경로 가져오기)
Function fnCallKakaoPayFileUrl(targetdate, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	sURL = KPay_Check_API_URL & "/files/v1/payments/history"

	'// 전송 테이터 Setting
	sParam = "target_date=" & targetdate								       'target_date(yyyyMMdd)
	sParam = sParam &"&bucket_id=" & KPay_PaymentsFileBID                  '파일서비스에서 제공하는 bucket_id
    
	sURL = sURL + "?" + sParam

	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "GET", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "PG_BIZAPI_KEY " & KPay_FileKey
	oXML.send		'전송
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제내역file 인증 통신 오류[001]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제내역file 인증 내부 오류[002]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPayFileUrl = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.url<>"" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayFileUrl = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제내역file 인증 오류[003]. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제내역file 인증 결과 파징 오류[004]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPayFileUrl = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// 결제내역 API 호출 함수(결제내역 Json 가져오기)
Function fnCallKakaoPayCheckList(targeturl, byref iStatus)
	dim oXML, sURL
	dim jsResult, oResult
	
	sURL = targeturl

	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "GET", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "PG_BIZAPI_KEY " & KPay_FileKey
	oXML.send		'전송
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제내역 인증 통신 오류[005]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제내역 인증 내부 오류[006]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPayCheckList = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.type<>"" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayCheckList = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제내역 인증 오류[007]. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "결제내역 인증 결과 파징 오류[008]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPayCheckList = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// 정산내역 API 호출 함수(파일 경로 가져오기)
Function fnCallKakaoPaySettlementsFileUrl(targetdate, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	sURL = KPay_Check_API_URL & "/files/v1/settlements/history"

	'// 전송 테이터 Setting
	sParam = "target_date=" & targetdate								       'target_date(yyyyMMdd)
	sParam = sParam &"&bucket_id=" & KPay_SettlementsFileBID                '파일서비스에서 제공하는 bucket_id
    
	sURL = sURL + "?" + sParam

	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "GET", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "PG_BIZAPI_KEY " & KPay_FileKey
	oXML.send		'전송
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "정산내역file 인증 통신 오류[001]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "정산내역file 인증 내부 오류[002]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPaySettlementsFileUrl = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.url<>"" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPaySettlementsFileUrl = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "정산내역file 인증 오류[003]. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "정산내역file 인증 결과 파징 오류[004]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPaySettlementsFileUrl = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// 정산내역 API 호출 함수(결제내역 Json 가져오기)
Function fnCallKakaoPaySettlementsCheckList(targeturl, byref iStatus)
	dim oXML, sURL
	dim jsResult, oResult
	
	sURL = targeturl

	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'결과객체 생성

	'// 호출 처리
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP컨퍼넌트 선언
	oXML.open "GET", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "PG_BIZAPI_KEY " & KPay_FileKey
	oXML.send		'전송
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "정산내역 인증 통신 오류[005]"
	end if
	Set oXML = Nothing	'컨퍼넌트 해제

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "정산내역 인증 내부 오류[006]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPaySettlementsCheckList = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// 결과값 Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.type<>"" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPaySettlementsCheckList = oResult			'파징된 정보 반환
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "정산내역 인증 오류[007]. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "정산내역 인증 결과 파징 오류[008]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPaySettlementsCheckList = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'==================================================
'// 결과 반환용 객체 선언
Class cResultKakao
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