<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
'+-------------------------------------------------------------------------------------------------------------+
'|                                  ī ī �� �� ��   �� ��   �� �� �� ��                                       |
'+----------------------------------------------+--------------------------------------------------------------+
'|                �� �� ��                      |                          ��    ��                            |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallChaiPayCancel                          | ��ü ��� ��û ȣ��                                          |
'| (������ȣ,��ұݾ�)                           | - ��ȯ�� : ��ҳ��� Object (������ code: ERR)               |
'+----------------------------------------------+--------------------------------------------------------------+
'| fnCallChaiPayPartialCancel                    | �κ� ��� ��û ȣ��                                          |
'| (������ȣ, �����ܾ�, ��ұݾ�)                | - ��ȯ�� : ��ҳ��� Object (������ code: ERR)               |
'+----------------------------------------------+--------------------------------------------------------------+

Dim ChaiPay_API_URL		        	    '// �鿣�� API ȣ�� URL
Dim ChaiPay_Check_API_URL				'// �������� API ȣ�� URL
Dim ChaiPay_PublicKey                    '// Public-API-Key (����Ű)
Dim ChaiPay_PrivateKey                    '// Private-API-Key (���Ű)

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
'// ��ü ��� ��û ȣ�� �Լ�
Function fnCallChaiPayCancel(npId, idempotencyKey, cancelAmount, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult
	
	sURL = ChaiPay_API_URL & "/v1/payment/" & npId & "/cancel"

	'// ���� ������ Setting
	sParam = "cancelAmount=" & cancelAmount										         '���� ������ȣ. 20��.
    
	Dim oRstMsg
	Set oRstMsg = new cResultChai											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Private-API-Key", ChaiPay_PrivateKey
    oXML.setRequestHeader "Idempotency-Key", idempotencyKey
	oXML.send sParam	'�Ķ���� ����
	response.write oXML.responseText & "<br>"
	response.end
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
		set fnCallChaiPayCancel = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.status="canceled" then
		oRstMsg.code = "Success"
		Set fnCallChaiPayCancel = oResult			'��¡�� ���� ��ȯ
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
		set fnCallChaiPayCancel = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'=================================================
'// ��ҿ�û ȣ�� �Լ�
Function fnCallChaiPayPartialCancel(npId, idempotencyKey, mainpaymentorg ,cancelAmount, byref iStatus)
	dim oXML, sURL, sParam
	dim jsResult, oResult

	sURL = ChaiPay_API_URL & "/v1/payment/" & npId & "/cancel"

	'// ���� ������ Setting
	sParam = "cancelAmount=" & cancelAmount					'��ұݾ�
	sParam = sParam &"&checkoutAmount=" & mainpaymentorg  '��ҿ�û���� �����ܾ�(�ߺ���� ������)

	Dim oRstMsg
	Set oRstMsg = new cResultChai											'�����ü ����

	'// ȣ�� ó��
	on Error Resume Next

	set oXML = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")	'xmlHTTP���۳�Ʈ ����
	oXML.open "POST", sURL, false
	oXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
	oXML.setRequestHeader "Private-API-Key", ChaiPay_PrivateKey
    oXML.setRequestHeader "Idempotency-Key", idempotencyKey
	oXML.send sParam	'�Ķ���� ����
'response.write mainpaymentorg & "<br>"
response.write oXML.status & "<br>"
response.write oXML.responseText & "<br>"
'response.end
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
		set fnCallChaiPayPartialCancel = oRstMsg
		Exit Function
	end if

	on Error Resume Next
	'// ����� Parsing
	set oResult = JSON.parse(jsResult)
	
	if oResult.status="canceled" then
		oRstMsg.code = "Success"
		Set fnCallChaiPayPartialCancel = oResult			'��¡�� ���� ��ȯ
	elseif oResult.status="confirmed" then
		oRstMsg.code = "Success"
		Set fnCallChaiPayPartialCancel = oResult			'��¡�� ���� ��ȯ
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
		set fnCallChaiPayPartialCancel = oRstMsg
	end if
	Set oRstMsg = Nothing
End Function

'==================================================
'// ��� ��ȯ�� ��ü ����
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