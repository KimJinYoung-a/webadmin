<%
Dim homeplusAPIURL, strInterface, homeplusVenderID, homepluspasswd
IF application("Svr_Info") = "Dev" THEN
	homeplusAPIURL = "http://112.108.7.201:7006/services/API2?wsdl"
	strInterface = "http://112.108.7.201:7006/api/services/API2"
	homeplusVenderID = "292811"
	homepluspasswd = "qwer1234"
Else
	homeplusAPIURL = "http://api.direct.homeplus.co.kr:17004?wsdl"
	strInterface = "http://api.direct.homeplus.co.kr:17004/api/services/API2"
	homeplusVenderID = "292811"
	homepluspasswd = "cube1010!!"
End if
'// Ȩ�÷��� �α��� Ȯ��(���ø����̼Ǻ����� ����)
If Application("homeplusAuthDate") = "" or Datediff("d", Application("homeplusAuthDate"), date()) > 0 Then
	Dim objXML, xmlDOM, strRst
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		strRst = ""
		strRst = strRst & "<?xml version=""1.0"" encoding=""utf-8""?>"
		strRst = strRst & "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/""  xmlns:SOAP-ENC=""http://schemas.xmlsoap.org/soap/encoding/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">"
		strRst = strRst & "    <SOAP-ENV:Body>"
		strRst = strRst & "        <m:login xmlns:m=""" & strInterface & """>"
		strRst = strRst & "            <venderId>"&homeplusVenderID&"</venderId>"
		strRst = strRst & "            <passwd>"&homepluspasswd&"</passwd>"
		strRst = strRst & "        </m:login>"
		strRst = strRst & "    </SOAP-ENV:Body>"
		strRst = strRst & "</SOAP-ENV:Envelope>"
		objXML.open "post", "" & homeplusAPIURL & "", False
		objXML.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
		objXML.setRequestHeader "Content-Length", Len(strRst)
		objXML.setRequestHeader "SOAPAction", strInterface & "#login"
		objXML.send(strRst)

		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				On Error Resume Next
					Application("loginCode") = xmlDOM.getElementsByTagName("ns1:code").item(0).text		'�α����ڵ� ����
					If (Application("loginCode") = "") OR (Application("loginCode") <> "E0000") Then	'�α��� ���е��� ���
						Response.Write "<script language=javascript>alert('Ȩ�÷��� �α��ο� �����Ͽ����ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
						Response.End
					End If

					Application("homeplusAuthDate") = now()												'�α��� �ð� ���
					If Err <> 0 then
						Response.Write "<script language=javascript>alert('Ȩ�÷��� �α��ο� �����Ͽ����ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
						Response.End
					End If
				On Error Goto 0
			Set xmlDOM = nothing
		Else
			Response.Write "<script language=javascript>alert('Ȩ�÷��� �α��ο� ������ �߻��߽��ϴ�.\n���߿� �ٽ� �õ��غ�����');history.back();</script>"
			Response.End
		End If
	Set objXML = nothing
End If
rw "�α�ð�: "& Application("homeplusAuthDate")
rw "����ð�: "&now()
%>