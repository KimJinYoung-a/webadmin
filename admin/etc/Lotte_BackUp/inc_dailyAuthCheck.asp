<%
	''//https://partner.lotte.com/main/Login.lotte 로그인정책과 같음 // 이곳도 같이 변경해야 하는듯.
	'// 롯데닷컴API 연동서버 URL
	dim lotteAPIURL, lotteAuthNo, lottenTenID, tenBrandCd, tenDlvCd, CertPasswd
	IF application("Svr_Info")="Dev" THEN
		'lotteAPIURL = "http://openapidev.lotte.com"	'' 테스트서버
		lotteAPIURL = "http://openapitest.lotte.com"	'' 테스트서버
		tenBrandCd = "14846"	'텐바(임시)
		tenDlvCd = "513564"		'배송정책코드
		CertPasswd = "1234"		'Dev는 비번 : 1234
	Else
		lotteAPIURL = "https://openapi.lotte.com"		'' 실서버
		tenBrandCd = "155112"	'텐바이텐
		tenDlvCd = "513484"
		CertPasswd = "cube101010"
	End if
	lottenTenID = "124072"					'텐바이텐ID

	'// 롯데닷컴 인증코드 확인(매일 업데이트; 어플리케이션변수에 저장)
	'Application("lotteAuthDate")="2012-01-01"

	if Application("lotteAuthDate")="" or datediff("d",Application("lotteAuthDate"),date())>0 then
		dim objXML, xmlDOM
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", lotteAPIURL & "/openapi/createCertification.lotte?strUserId=" & lottenTenID & "&strPassWd="&CertPasswd&"", false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then

			'//전달받은 내용 확인
			'Response.contentType = "text/xml; charset=euc-kr"
			'response.write BinaryToText(objXML.ResponseBody, "euc-kr")
			'response.End

			'XML을 담을 DOM 객체 생성
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
			xmlDOM.async = False
			'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
			xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

			on Error Resume Next
				Application("lotteAuthNo") = xmlDOM.getElementsByTagName("SubscriptionId").item(0).text		'인증번호 저장
				Application("lotteAuthDate") = now()			'인증시간 기록
				if Err<>0 then
					Response.Write "<script language=javascript>alert('Lotte.com인증에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
					Response.End
				end if
			on Error Goto 0

			Set xmlDOM = Nothing
			
			dim iisql
			iisql = "update db_etcmall.dbo.tbl_outmall_ini "&VbCRLF
			iisql = iisql & " set iniVal='"&Application("lotteAuthNo")&"'"&VbCRLF
			iisql = iisql & " ,lastupdate=getdate()"&VbCRLF
			iisql = iisql & " where mallid='lotteCom'"&VbCRLF
			iisql = iisql & " and inikey='auth'"
			dbget.Execute iisql
		else
			Response.Write "<script language=javascript>alert('Lotte.com과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');history.back();</script>"
			Response.End
		end if
		Set objXML = Nothing

	end if

	lotteAuthNo = Application("lotteAuthNo")
%>