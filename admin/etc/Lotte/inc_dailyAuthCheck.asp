<%
''//https://partner.lotte.com/main/Login.lotte �α�����å�� ���� // �̰��� ���� �����ؾ� �ϴµ�.
'// �Ե�����API �������� URL
dim lotteAPIURL, lotteAuthNo, lottenTenID, tenBrandCd, tenDlvCd, CertPasswd
IF application("Svr_Info")="Dev" THEN
	'lotteAPIURL = "http://openapidev.lotte.com"	'' �׽�Ʈ����
	lotteAPIURL = "http://openapitest.lotte.com"	'' �׽�Ʈ����
	tenBrandCd = "14846"	'�ٹ�(�ӽ�)
	tenDlvCd = "513564"		'�����å�ڵ�
	CertPasswd = "1234"		'Dev�� ��� : 1234
Else
	lotteAPIURL = "https://openapi.lotte.com"		'' �Ǽ���
	tenBrandCd = "155112"	'�ٹ�����
	tenDlvCd = "513484"
	CertPasswd = "store101010**"
End if
lottenTenID = "124072"					'�ٹ�����ID

Dim updateAuth, dbAuthNo
dim iisql, tmplotteComAuthNo
iisql = "select top 1 isnull(iniVal, '') as iniVal, lastupdate "&VbCRLF
iisql = iisql & " from db_etcmall.dbo.tbl_outmall_ini"&VbCRLF
iisql = iisql & " where mallid='lotteCom'"&VbCRLF
iisql = iisql & " and inikey='auth'"
rsget.Open iisql, dbget, 1
if not rsget.Eof then
    dbAuthNo	= rsget("iniVal")
    updateAuth	= rsget("lastupdate")
end if
rsget.close

'// �Ե����� �����ڵ� Ȯ��(���� ������Ʈ; ���ø����̼Ǻ����� ����)
'Application("lotteAuthDate")="2012-01-01"
'2015-06-08 ������ �ϴ� ����Ű �Ϸ翡�� 12�ð����� ����
'2015-07-13 ������ �ϴ� ����Ű �Ϸ翡�� 6�ð����� ����
'	if Application("lotteAuthDate")="" or datediff("h",Application("lotteAuthDate"),now())>6 then
If DateDiff("h", updateAuth, now()) > 12 OR dbAuthNo = "" then
	dim objXML, xmlDOM
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/createCertification.lotte?strUserId=" & lottenTenID & "&strPassWd="&CertPasswd&"", false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then
		'//���޹��� ���� Ȯ��
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
		'response.End

		'XML�� ���� DOM ��ü ����
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM ��ü�� XML�� ��´�.(���̳ʸ� �����ͷ� �޾Ƽ� euc-kr�� ��ȯ(�ѱ۹���))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

		on Error Resume Next
			tmplotteComAuthNo = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text		'������ȣ ����
			if Err<>0 then
				Response.Write "<script language=javascript>alert('Lotte.com������ ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
				Response.End
			end if
		on Error Goto 0

		Set xmlDOM = Nothing

		iisql = "update db_etcmall.dbo.tbl_outmall_ini "&VbCRLF
		iisql = iisql & " set iniVal = '"&tmplotteComAuthNo&"'"&VbCRLF
		iisql = iisql & " ,lastupdate = getdate()"&VbCRLF
		iisql = iisql & " where mallid = 'lotteCom'"&VbCRLF
		iisql = iisql & " and inikey='auth'"
		dbget.Execute iisql
	else
		Response.Write "<script language=javascript>alert('Lotte.com�� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
		Response.End
	end if
	Set objXML = Nothing
'	end if
	lotteAuthNo = tmplotteComAuthNo
Else
	lotteAuthNo = dbAuthNo
End If
%>