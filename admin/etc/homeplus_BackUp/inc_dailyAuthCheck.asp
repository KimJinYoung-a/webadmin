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
'// 홈플러스 로그인 확인(어플리케이션변수에 저장)
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
					Application("loginCode") = xmlDOM.getElementsByTagName("ns1:code").item(0).text		'로그인코드 저장
					If (Application("loginCode") = "") OR (Application("loginCode") <> "E0000") Then	'로그인 실패됐을 경우
						Response.Write "<script language=javascript>alert('홈플러스 로그인에 실패하였습니다.\n나중에 다시 시도해보세요');history.back();</script>"
						Response.End
					End If

					Application("homeplusAuthDate") = now()												'로그인 시간 기록
					If Err <> 0 then
						Response.Write "<script language=javascript>alert('홈플러스 로그인에 실패하였습니다.\n나중에 다시 시도해보세요');history.back();</script>"
						Response.End
					End If
				On Error Goto 0
			Set xmlDOM = nothing
		Else
			Response.Write "<script language=javascript>alert('홈플러스 로그인에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
			Response.End
		End If
	Set objXML = nothing
End If
rw "로긴시간: "& Application("homeplusAuthDate")
rw "현재시간: "&now()
%>