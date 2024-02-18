<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// 변수선언
	dim MDCode, MDName, SellFeeType, NormalSellFee, EventSellFee
	dim strSql, actCnt

	actCnt = 0		'실갱신건수

	'// 롯데닷컴 담당MD 조회
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/searchMDListOpenApi.lotte?subscriptionId=" & lotteAuthNo, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
	If objXML.Status = "200" Then

		'//전달받은 내용 확인
		'Response.contentType = "text/xml; charset=euc-kr"
		'response.write BinaryToText(objXML.ResponseBody, "euc-kr")

		'XML을 담을 DOM 객체 생성
		Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		'DOM 객체에 XML을 담는다.(바이너리 데이터로 받아서 euc-kr로 변환(한글문제))
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
		
		on Error Resume Next
			MDCnt = xmlDOM.getElementsByTagName("MDCount").item(0).text		'담당MD수
			if Err<>0 then
				Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
				Response.End
			end if

			if MDCnt>0 then
				'// 트랜젝션 시작
				dbget.beginTrans

				'모든 MD사용여부 변경
				strSql = "update db_temp.dbo.tbl_lotte_MDInfo Set isUsing='N', lastupdate=getdate() Where isUsing='Y' "
				dbget.Execute(strSql)

				'// MDInfo Loop
				Set MDInfo = xmlDOM.getElementsByTagName("MDInfo")
				for each SubNodes in MDInfo
					MDCode			= Trim(SubNodes.getElementsByTagName("MDCode").item(0).text)		'담당MD코드
					MDName			= Trim(SubNodes.getElementsByTagName("MDName").item(0).text)		'담당MD명
					SellFeeType		= Trim(SubNodes.getElementsByTagName("SellFeeType").item(0).text)	'마진구분
					NormalSellFee	= Trim(SubNodes.getElementsByTagName("NormalSellFee").item(0).text)	'정상수수료
					EventSellFee	= Trim(SubNodes.getElementsByTagName("EventSellFee").item(0).text)	'행사수수료

					'MD존재여부 확인
					strSql = "Select count(MDCode) From db_temp.dbo.tbl_lotte_MDInfo Where MDCode='" & MDCode & "'"
					rsget.Open strSql,dbget,1

					if rsget(0)>0 then
						'// 존재 -> 사용함
						strSql = "update db_temp.dbo.tbl_lotte_MDInfo Set isUsing='Y' Where MDCode='" & MDCode & "'"
						dbget.Execute(strSql)
					else
						'// 없음 -> 신규등록
						strSql = "Insert into db_temp.dbo.tbl_lotte_MDInfo (MDCode, MDName, SellFeeType, NormalSellFee, EventSellFee) values " &_
								" ('" & MDCode & "'" &_
								", '" & html2db(MDName) & "'" &_
								", '" & SellFeeType & "'" &_
								", '" & NormalSellFee & "'" &_
								", '" & EventSellFee & "')"
						dbget.Execute(strSql)
						actCnt = actCnt+1
					end if

					rsget.Close
				Next
				Set MDInfo = Nothing

				'##### DB 저장 처리 #####
			    If Err.Number = 0 Then
			    	dbget.CommitTrans				'커밋(정상)
			    	Response.Write "<script language=javascript>alert('" & actCnt & "건이 정상적으로 갱신되었습니다.');parent.history.go(0);</script>"
			    Else
			        dbget.RollBackTrans				'롤백(에러발생시)
			        Response.Write "<script language=javascript>alert('자료 저장 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
			    End If
			else
				Response.Write "<script language=javascript>alert('롯데닷컴에 지정되어있는 담당MD가 없습니다.\n롯데닷컴 담당자에게 문의해주세요.');</script>"
				Response.End
			end if
		on Error Goto 0

		Set xmlDOM = Nothing
	else
		Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
		Response.End
	end if
	Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->