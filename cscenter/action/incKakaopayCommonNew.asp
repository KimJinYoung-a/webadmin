<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'+-------------------------------------------------------------------------------------------------------------+
'|                                  ī ī �� �� ��   �� ��   �� �� �� ��                                       |
'+----------------------------------------------+--------------------------------------------------------------+
'|                �� �� ��                      |                          ��    ��                            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPayCancel                        | ��ü ��� ��û ȣ��                                          |
'| (������ȣ,��ұݾ�)                           | - ��ȯ�� : ��ҳ��� Object (������ code: ERR)               |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPayPartialCancel                  | �κ� ��� ��û ȣ��                                          |
'| (������ȣ,��ұݾ�)                           | - ��ȯ�� : ��ҳ��� Object (������ code: ERR)               |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPayFileUrl  		                 | ���� ���� ���� url ��û ȣ��                                  |
'| (��ȸ��¥)              				         | - ��ȯ�� : �������� ���� ��� (������ code: ERR)            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPayCheckList                      | ���� ���� ��û ȣ��                                          |
'| (����URL)              				        | - ��ȯ�� : �������� Json (������ code: ERR)                 |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPaySettlementsFileUrl              | ���� ���� ��û ȣ��                                          |
'| (��ȸ��¥)              				        | - ��ȯ�� : ���� ���� ���� ��� (������ code: ERR)            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallKakaoPaySettlementsCheckList           | �������� ��û ȣ��                                           |
'| (����URL)              				        | - ��ȯ�� : ���� ���� Json (������ code: ERR)                |
'+----------------------------------------------+--------------------------------------------------------------+

Dim KPay_API_URL		        	  '// �鿣�� API ȣ�� URL
Dim KPay_Check_API_URL				'// �������� API ȣ�� URL
Dim KPay_PartnerID                     '// īī������ ��Ʈ�� ID
Dim KPay_ClientId, KPay_ClientKey      '// �ٹ����� ���� ID �� ��ũ��Ű
Dim KPay_FileKey					   '//���系��, ���� ���� ����
Dim KPay_PaymentsFileBID  			  '//���系�� BID(bucket_id)
Dim KPay_SettlementsFileBID			  '//���곻�� BID(bucket_id)

'īī������ �⺻ ���
KPay_API_URL = "https://kapi.kakao.com"
KPay_Check_API_URL = "https://biz-api.kakaopay.com"
if (application("Svr_Info")="Dev") then
	'����Ű
	KPay_PartnerID = "TC0ONETIME"							 '������ �ڵ�. 10��
	KPay_ClientId = "cb8b50980734335b667cffb32781e5a1"      'Admin Key
	KPay_ClientKey = "d0e01cbc77f59fa73969f46ae83dd9ca"     'RestAPI Key
	KPay_PaymentsFileBID = "B656960246"
	KPay_SettlementsFileBID = "B997240247"
    KPay_FileKey = "cc061a13c0e1f1580c8c4168439b916d585c61ffb219cd65d8cc2cbfd7f74766" '���� ���� Ű

else
	'����Ű
	KPay_PartnerID = "C371930065"							 '������ �ڵ�. 10��
	KPay_ClientId = "5a1f82cd75e1002b529edc6f213d875a"      'Admin Key
	KPay_ClientKey = "b4e7e01a2ade8ecedc5c6944941ffbd4"     'RestAPI Key
	KPay_PaymentsFileBID = "B656960246"
	KPay_SettlementsFileBID = "B997240247"
    KPay_FileKey = "cc061a13c0e1f1580c8c4168439b916d585c61ffb219cd65d8cc2cbfd7f74766"
end if

'=================================================
'// ��ҿ�û ȣ�� �Լ�
Function fnCallKakaoPayCancel(npId, cancelAmount, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	sURL = KPay_API_URL & "/v1/payment/cancel"

	'// ���� ������ Setting
	sParam = "tid=" & npId										         '���� ������ȣ. 20��.
	sParam = sParam &"&cid=" & KPay_PartnerID                          '������ �ڵ�. 10��
    sParam = sParam &"&cancel_amount=" & cancelAmount                '��ұݾ�
    sParam = sParam &"&cancel_tax_free_amount=" & 0        				'��� ����� �ݾ�
    
	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "KakaoAK " & KPay_ClientId
	oXML.send sParam	'�Ķ���� ����
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "����������� ��� ����[004]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "����������� ���� ����[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPayCancel = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.status="CANCEL_PAYMENT" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayCancel = oResult			'��¡�� ���� ��ȯ
	elseif oResult.status="PART_CANCEL_PAYMENT" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayCancel = oResult			'��¡�� ���� ��ȯ
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "����������� ����. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "������� ������� ��¡ ����[006]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPayCancel = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// ��ҿ�û ȣ�� �Լ�
Function fnCallKakaoPayPartialCancel(npId, mainpaymentorg ,cancelAmount, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	sURL = KPay_API_URL & "/v1/payment/cancel"

	'// ���� ������ Setting
	sParam = "tid=" & npId										        '���� ������ȣ. 20��.
	sParam = sParam &"&cid=" & KPay_PartnerID                         '������ �ڵ�. 10��
    sParam = sParam &"&cancel_amount=" & cancelAmount               '��ұݾ�
    sParam = sParam &"&cancel_tax_free_amount=" & 0        			  '��� ����� �ݾ�
	sParam = sParam &"&cancel_available_amount=" & mainpaymentorg   '���� �����ݾ�
    
	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "KakaoAK " & KPay_ClientId
	oXML.send sParam	'�Ķ���� ����
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "����������� ��� ����[004]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "����������� ���� ����[005]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPayPartialCancel = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.status="CANCEL_PAYMENT" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayPartialCancel = oResult			'��¡�� ���� ��ȯ
	elseif oResult.status="PART_CANCEL_PAYMENT" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayPartialCancel = oResult			'��¡�� ���� ��ȯ
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "����������� ����. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "������� ������� ��¡ ����[006]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPayPartialCancel = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// �������� API ȣ�� �Լ�(���� ��� ��������)
Function fnCallKakaoPayFileUrl(targetdate, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	sURL = KPay_Check_API_URL & "/files/v1/payments/history"

	'// ���� ������ Setting
	sParam = "target_date=" & targetdate								       'target_date(yyyyMMdd)
	sParam = sParam &"&bucket_id=" & KPay_PaymentsFileBID                  '���ϼ��񽺿��� �����ϴ� bucket_id
    
	sURL = sURL + "?" + sParam

	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "GET", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "PG_BIZAPI_KEY " & KPay_FileKey
	oXML.send		'����
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "��������file ���� ��� ����[001]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "��������file ���� ���� ����[002]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPayFileUrl = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.url<>"" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayFileUrl = oResult			'��¡�� ���� ��ȯ
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "��������file ���� ����[003]. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "��������file ���� ��� ��¡ ����[004]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPayFileUrl = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// �������� API ȣ�� �Լ�(�������� Json ��������)
Function fnCallKakaoPayCheckList(targeturl, byref iStatus)
	dim oXML, sURL
	dim jsResult, oResult
	
	sURL = targeturl

	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "GET", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "PG_BIZAPI_KEY " & KPay_FileKey
	oXML.send		'����
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ���� ��� ����[005]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ���� ���� ����[006]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPayCheckList = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.type<>"" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPayCheckList = oResult			'��¡�� ���� ��ȯ
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ���� ����[007]. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "�������� ���� ��� ��¡ ����[008]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPayCheckList = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// ���곻�� API ȣ�� �Լ�(���� ��� ��������)
Function fnCallKakaoPaySettlementsFileUrl(targetdate, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	sURL = KPay_Check_API_URL & "/files/v1/settlements/history"

	'// ���� ������ Setting
	sParam = "target_date=" & targetdate								       'target_date(yyyyMMdd)
	sParam = sParam &"&bucket_id=" & KPay_SettlementsFileBID                '���ϼ��񽺿��� �����ϴ� bucket_id
    
	sURL = sURL + "?" + sParam

	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "GET", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "PG_BIZAPI_KEY " & KPay_FileKey
	oXML.send		'����
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "���곻��file ���� ��� ����[001]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "���곻��file ���� ���� ����[002]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPaySettlementsFileUrl = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.url<>"" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPaySettlementsFileUrl = oResult			'��¡�� ���� ��ȯ
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "���곻��file ���� ����[003]. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "���곻��file ���� ��� ��¡ ����[004]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPaySettlementsFileUrl = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// ���곻�� API ȣ�� �Լ�(�������� Json ��������)
Function fnCallKakaoPaySettlementsCheckList(targeturl, byref iStatus)
	dim oXML, sURL
	dim jsResult, oResult
	
	sURL = targeturl

	Dim oRstMsg
	Set oRstMsg = new cResultKakao											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "GET", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Authorization", "PG_BIZAPI_KEY " & KPay_FileKey
	oXML.send		'����
	'response.write oXML.responseText & "<br>"
	iStatus=oXML.status
	if oXML.status=200 then
		jsResult = oXML.responseText
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "���곻�� ���� ��� ����[005]"
	end if
	Set oXML = Nothing	'���۳�Ʈ ����

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "���곻�� ���� ���� ����[006]"
	end if

	On Error Goto 0

	if oRstMsg.code="ERR" then
		set fnCallKakaoPaySettlementsCheckList = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.type<>"" then
		oRstMsg.code = "Success"
		Set fnCallKakaoPaySettlementsCheckList = oResult			'��¡�� ���� ��ȯ
	else
		oRstMsg.code = "ERR"
		oRstMsg.message = "���곻�� ���� ����[007]. " & oResult.msg
	end if
	set oResult = Nothing

	IF (Err) then
		oRstMsg.code = "ERR"
		oRstMsg.message = "���곻�� ���� ��� ��¡ ����[008]"
	end if

	On Error Goto 0

    if oRstMsg.code="ERR" then
		set fnCallKakaoPaySettlementsCheckList = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'==================================================
'// ��� ��ȯ�� ��ü ����
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