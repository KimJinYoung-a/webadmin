<%
'// �Ե����̸�API �������� URL
Dim ltiMallAPIURL, ltiMallAuthNo, ltiMallTenID, tenBrandCd, tenDlvCd, tenDlvFreeCd
IF application("Svr_Info") = "Dev" THEN
	'ltiMallAPIURL = "http://openapidev.lotteimall.com"	'' �׽�Ʈ����
	ltiMallAPIURL = "http://openapitst.lotteimall.com"	'' �׽�Ʈ����
	tenDlvCd = "26645"
	tenDlvFreeCd = "577045"
Else
	ltiMallAPIURL = "https://openapi.lotteimall.com"		'' �Ǽ���
	tenDlvCd = "23725" 
	tenDlvFreeCd = "577045"
End if
ltiMallTenID = "011799LT"
'// �Ե����̸� �����ڵ� Ȯ��(���� ������Ʈ; ���ø����̼Ǻ����� ����)
'Application("ltiMallAuthDate")="2012-01-01"
If Application("ltiMallAuthDate") = "" or Datediff("d", Application("ltiMallAuthDate"), date()) > 0 Then
	Dim objXML, xmlDOM
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	    objXML.Open "GET", ltiMallAPIURL & "/openapi/createCertification.lotte?strUserId=" & ltiMallTenID & "&strPassWd=cube101010!*", False
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
			On Error Resume Next
				Application("ltiMallAuthNo") = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text		'������ȣ ����
				If Application("ltiMallAuthNo") = "" Then														'���� ���е��� ���
					Response.Write "<script language=javascript>alert('Lotteimall.com�����ڵ� ���ۿ����� �߻��Ͽ����ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
					Response.End					
				End If

				Application("ltiMallAuthDate") = now()															'�����ð� ���
				If Err <> 0 then
					Response.Write "<script language=javascript>alert('Lotteimall.com������ ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
					Response.End
				End If
			On Error Goto 0
			Set xmlDOM = Nothing
		Else
			Response.Write "<script language=javascript>alert('Lotteimall.com�� ����߿� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
			Response.End
		End If
	Set objXML = Nothing
End If
ltiMallAuthNo = Application("ltiMallAuthNo")
%>