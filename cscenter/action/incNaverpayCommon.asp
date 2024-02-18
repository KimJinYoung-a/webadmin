<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'+-------------------------------------------------------------------------------------------------------------+
'|                                  �� �� �� �� ��   �� ��   �� �� �� ��                                       |
'+----------------------------------------------+--------------------------------------------------------------+
'|                �� �� ��                      |                          ��    ��                            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayReserve(                       | ���� ���� ȣ��                                               |
'|        �ֹ���ȣ,��ǰ��,��ǰ��,�Ѱ����ݾ�     | - ��ȯ�� : �����ȣ (������ ERR)                             |
'|        ,�����ݾ�,��ۺ�,�ֹ����̸�)          |                                                              |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayApply(������ȣ)                | ���� ��û(����) ȣ��                                         |
'|                                              | - ��ȯ�� : ������ȣ Object (������ code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayCheck(������ȣ)                | ���� ���� ��û(Ȯ��) ȣ��                                    |
'|                                              | - ��ȯ�� : �������� Object (������ code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayCashAmt(������ȣ)              | ���� ������ ���� ��� �ݾ� ��ȸ ȣ��                         |
'|                                              | - ��ȯ�� : �������� Object (������ code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayReceipt(������ȣ)              | �ſ�ī�� ���� ��ǥ ��ȸ ȣ��                                 |
'|                                              | - ��ȯ�� : �ſ�ī�� ������ǥ ��ȸ URL (������ ERR)           |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayCancel(������ȣ,��ұݾ�,����,��û��) | ��ҿ�û ȣ��                                                |
'|                                              | - ��ȯ�� : ��ҳ��� Object (������ code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallNaverPayDlvFinish(������ȣ)            | �ŷ��Ϸ� ȣ��                                                |
'|                                              | - ��ȯ�� : ��ҳ��� Object (������ code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+



Dim NPay_API_URL			'// �鿣�� API ȣ�� URL
Dim NPay_SvcPC_URL			'// PC�� ���� API ȣ�� URL
Dim NPay_SvcMobile_URL		'// ����� ���� API ȣ�� URL
Dim NPay_PartnerID			'// ���̹����� ��Ʈ�� ID
Dim NPay_ClientId, NPay_ClientKey	'// �ٹ����� ���� ID �� ��ũ��Ű

'���̹����� �⺻ ���
if (application("Svr_Info")="Dev") then
	NPay_API_URL = "https://dev.apis.naver.com"
	NPay_SvcPC_URL = "https://alpha2-pay.naver.com"
	NPay_SvcMobile_URL = "https://alpha2-m.pay.naver.com"
else
	NPay_API_URL = "https://apis.naver.com"
	NPay_SvcPC_URL = "https://pay.naver.com"
	NPay_SvcMobile_URL = "https://m.pay.naver.com"
end if

'����Ű
NPay_PartnerID = "tenbyten"
NPay_ClientId = "FyDBW8XfYK4wly9KVYVz"
NPay_ClientKey = "AzeQydNElY"

'=================================================

'// �������� ȣ�� �Լ�
Function fnCallNaverPayReserve(ordno,itemnm,itemno,totprc,taxAmt,dlvprc,ordunm)
	dim oXML, sURL, sParam
	dim jsResult, oResult, errMsg
	
	'// �������� (https://[API������]/[������ID]/naverpay/payments/v1/reserve)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/reserve"

	'// ���� ������ Setting
	sParam = "modelVersion=2"										'���� ������� (1:��ý���, 2:�����Ľ���)
	sParam = sParam & "&merchantPayKey=" & ordno							'�ֹ���ȣ (�ӽ�)
	sParam = sParam & "&productName=" & server.URLEncode(itemnm)			'��ǥ ��ǰ�� (�ݵ�� 1��)
	sParam = sParam & "&productCount=" & itemno								'��ǰ����
	sParam = sParam & "&totalPayAmount=" & totprc							'�� ���� �ݾ�
	sParam = sParam & "&taxScopeAmount=" & taxAmt							'���� �ݾ� (����+�鼼=�Ѱ����ݾ�)
sParam = sParam & "&taxExScopeAmount=" & (totprc-taxAmt)				'�鼼 �ݾ� (0�̶� ����)
	if dlvprc>0 then
		sParam = sParam & "&deliveryFee=" & dlvprc			'��ۺ�
	end if
	sParam = sParam & "&returnUrl=" & server.URLEncode(SSLUrl & "/inipay/naverpay/naverPayResult.asp?ordsn=" & rdmSerialEnc(ordno))	'���� �Ϸ� �� �̵��� URL (�ӽ��ֹ���ȣ ��ȣȭ �� ����)
	''sParam = sParam & "&purchaserName=" & server.URLEncode(ordunm)			'������ ����(������)

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'�Ķ���� ����

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		errMsg = "ERR:�������� ��� ����[001]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		errMsg = "ERR:�������� ���� ����[002]"
	end if

	On Error Goto 0

	if errMsg<>"" then
		fnCallNaverPayReserve = errMsg
		Exit Function
	end if


	on Error Resume Next

	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.code="Success" then
		fnCallNaverPayReserve = oResult.body.reserveId
	else
		errMsg = "ERR:�������� ����. " & oResult.message
	end if
	
	set oResult = Nothing

	IF (Err) then
		errMsg = "ERR:������ ��¡ ����[003]"
	end if

	On Error Goto 0

	if errMsg<>"" then
		fnCallNaverPayReserve = errMsg
	end if

End Function


'// ������û ȣ�� �Լ�
Function fnCallNaverPayApply(npId)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	'// ������û (https://[API������]/[������ID]/naverpay/payments/v1/apply/payment)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/apply/payment"

	'// ���� ������ Setting
	sParam = "paymentId=" & npId										'���̹����� ���� ��ȣ

	Dim oRstMsg
	Set oRstMsg = new cResult											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'�Ķ���� ����

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ��� ����[004]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ���� ����[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayApply = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.code="Success" then
		oRstMsg.code = "Success"
		oRstMsg.message = oResult.body.paymentId
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ����. " & oResult.message
	end if
	
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "������� ��¡ ����[006]"
	end if

	On Error Goto 0

	'// ��� ��ȯ
	set fnCallNaverPayApply = oRstMsg
	Set oRstMsg = Nothing
End Function


'// ���� ���� ��ȸ ȣ�� �Լ�
Function fnCallNaverPayCheck(npId)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	'// ������û (https://[API������]/[������ID]/naverpay/payments/v1/list-of-payment)
	'// ���̹� ���� ���� ���� ��ȸ url ����(2020.12.8), v1/list-of-payment-->v2.2/list/history
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v2.2/list/history"

	'// ���� ������ Setting
	'sParam = "paymentId=" & npId										'���̹����� ���� ��ȣ
	'// ���� ������ ���� ����
	sParam = "{"
	sParam = sParam &"""paymentId"":"""&CStr(npId)&""""
	sParam = sParam &"}"	

	Dim oRstMsg
	Set oRstMsg = new cResult											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	'oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "Content-Type", "application/json"	
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'�Ķ���� ����

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ��� ����[007]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ���� ����[008]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCheck = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.code="Success" then
		oRstMsg.code = "Success"
		''fnCallNaverPayCheck = jsResult			'Raw ����� ��ȯ
		Set fnCallNaverPayCheck = oResult			'��¡�� ���� ��ȯ
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ����. " & oResult.message
	end if
	
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "���ΰ�� ��¡ ����[009]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCheck = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function



'// ���� ������ ���� ��� �ݾ� ��ȸ
Function fnCallNaverPayCashAmt(npId)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	'// ���ݼ� �����ݾ� ��û (https://[API������]/[������ID]/naverpay/payments/v1/receipt/cash-amount)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/receipt/cash-amount"

	'// ���� ������ Setting
	sParam = "paymentId=" & npId										'���̹����� ���� ��ȣ

	Dim oRstMsg
	Set oRstMsg = new cResult											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'�Ķ���� ����

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "���ݿ����� ��� ����[010]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "���ݿ����� ���� ����[011]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCashAmt = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.code="Success" then
		oRstMsg.code = "Success"
		''fnCallNaverPayCashAmt = jsResult			'Raw ����� ��ȯ
		Set fnCallNaverPayCashAmt = oResult			'��¡�� ���� ��ȯ
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "���ݿ����� ����. " & oResult.message
	end if
	
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "���ݴ�� Ȯ�ΰ�� ��¡ ����[012]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCashAmt = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function

'// �ſ�ī�� ���� ��ǥ ��ȸ ȣ�� �Լ�
Function fnCallNaverPayReceipt(npId)
	dim oXML, sURL, sParam
	dim jsResult, oResult, errMsg
	
	'// �ſ�ī�� ������ǥ (https://[API������]/[������ID]/naverpay/payments/v1/receipt/credit-card)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/receipt/credit-card"

	'// ���� ������ Setting
	sParam = "paymentId=" & npId										'���̹����� ���� ��ȣ

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'�Ķ���� ����

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		errMsg = "ERR:��ǥ��ȸ ��� ����[001]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		errMsg = "ERR:��ǥ��ȸ ���� ����[002]"
	end if

	On Error Goto 0

	if errMsg<>"" then
		fnCallNaverPayReceipt = errMsg
		Exit Function
	end if


	on Error Resume Next

	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.code="Success" then
		fnCallNaverPayReceipt = oResult.body.receiptUrl
	else
		errMsg = "ERR:��ǥ��ȸ ����. " & oResult.message
	end if
	
	set oResult = Nothing

	IF (Err) then
		errMsg = "ERR:��ǥ��ȸ ��¡ ����[003]"
	end if

	On Error Goto 0

	if errMsg<>"" then
		fnCallNaverPayReceipt = errMsg
	end if

End Function


'// ��ҿ�û ȣ�� �Լ�
Function fnCallNaverPayCancel(npId,cancelAmount,cancelReason,nPayCancelRequester)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	'// ������û (https://[API������]/[������ID]/naverpay/payments/v1/cancel)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/cancel"

	'// ���� ������ Setting
	sParam = "paymentId=" & npId										'���̹����� ���� ��ȣ
	''sParam = sParam &"&merchantPayKey=" & merchantPayKey                '�������� ������ȣ(����)
    sParam = sParam &"&cancelAmount=" & cancelAmount                    '��ұݾ�
    sParam = sParam &"&cancelReason=" & cancelReason                    '��һ���
    sParam = sParam &"&cancelRequester=" & nPayCancelRequester          '��ҿ�û��(1:������,2:������������)
    sParam = sParam &"&taxScopeAmount=" & cancelAmount                   '�������ݾ�(����)
    sParam = sParam &"&taxExScopeAmount=" & 0                            '�鼼���ݾ�(����)
    
	Dim oRstMsg
	Set oRstMsg = new cResult											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'�Ķ���� ����

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ��� ����[004]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ���� ����[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayCancel = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.code="Success" then
		oRstMsg.code = "Success"
		Set fnCallNaverPayCancel = oResult			'��¡�� ���� ��ȯ
		''oRstMsg.message = oResult.body.paymentId
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ����. " & oResult.message
	end if
	
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "������� ��¡ ����[006]"
	end if

	On Error Goto 0
    
    if oRstMsg.code="ERR" then
		set fnCallNaverPayCancel = oRstMsg
	end if

	Set oRstMsg = Nothing
	
End Function

'// �ŷ��Ϸ� ȣ�� �Լ�
Function fnCallNaverPayDlvFinish(npId)
    dim oXML, sURL, sParam
	dim jsResult, oResult
	
	'// ������û (https://[API������]/[������ID]/naverpay/payments/v1/purchase-confirm)
	sURL = NPay_API_URL & "/" & NPay_PartnerID & "/naverpay/payments/v1/purchase-confirm"

	'// ���� ������ Setting
	sParam = "paymentId=" & npId										'���̹����� ���� ��ȣ
    sParam = sParam &"&requester=2"                                     '��û��(1:������,2:������������)
    
	Dim oRstMsg
	Set oRstMsg = new cResult											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	oXML.setRequestHeader "X-Naver-Client-Id", NPay_ClientId
	oXML.setRequestHeader "X-Naver-Client-Secret", NPay_ClientKey
	oXML.send sParam	'�Ķ���� ����

	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ��� ����[004]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ���� ����[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallNaverPayDlvFinish = oRstMsg
		Exit Function
	end if


	on Error Resume Next

	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.code="Success" then
		oRstMsg.code = "Success"
		Set fnCallNaverPayDlvFinish = oResult			'��¡�� ���� ��ȯ
		''oRstMsg.message = oResult.body.paymentId
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ����. " & oResult.message
	end if
	
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "������� ��¡ ����[006]"
	end if

	On Error Goto 0
    
    if oRstMsg.code="ERR" then
		set fnCallNaverPayDlvFinish = oRstMsg
	end if

	Set oRstMsg = Nothing
End Function

'==================================================
'// ��� ��ȯ�� ��ü ����
Class cResult
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