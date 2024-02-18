<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// 변수선언
	dim SGrpCnt, GrpCnt, lp, SGrpInfo, GrpInfo
	dim MDCode, groupCode, SuperGroupName, GroupName
	dim strSql, actCnt

	actCnt = 0		'실갱신건수

	MDCode = Request("mdcd")

	'// 롯데닷컴 MD상품군 조회
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", lotteAPIURL & "/openapi/searchMDGsgrListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&md_id=" & MDCode, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()
rw lotteAPIURL & "/openapi/searchMDGsgrListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&md_id=" & MDCode
'response.end

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
		
		'on Error Resume Next
			SGrpCnt = xmlDOM.getElementsByTagName("SuperGroupCount").item(0).text		'상위상품군 카운트
			if Err<>0 then
				Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
				Response.End
			end if
rw "SGrpCnt="&SGrpCnt
			if SGrpCnt>0 then
				'// 트랜젝션 시작
				dbget.beginTrans

				'모든 상품군사용여부 변경
				strSql = "update db_temp.dbo.tbl_lotte_MDCateGrp Set isUsing='N', lastupdate=getdate() Where isUsing='Y' and MDCode='"&MDCode&"'"
				''rw strSql
				dbget.Execute(strSql)

				'// SGrpInfo Loop
				Set SGrpInfo = xmlDOM.getElementsByTagName("SuperGroupInfo")
				for each SGNodes in SGrpInfo
					SuperGroupName	= Trim(SGNodes.getElementsByTagName("SuperGroupName").item(0).text)		'상위상품군명
					GrpCnt			= SGNodes.getElementsByTagName("SubGroupCount").item(0).text			'하위상품군수

					if GrpCnt>0 then
						Set GrpInfo = SGNodes.getElementsByTagName("SubGroupInfo")
						for each SubNodes in GrpInfo
							groupCode	= Trim(SubNodes.getElementsByTagName("GroupCode").item(0).text)		'그룹코드
							GroupName	= Trim(SubNodes.getElementsByTagName("GroupName").item(0).text)		'그룹명
	
							'상품군존재여부 확인
							strSql = "Select count(*) From db_temp.dbo.tbl_lotte_MDCateGrp Where groupCode='" & groupCode & "' and MDCode='" & MDCode & "'"
							rsget.Open strSql,dbget,1
		
							if rsget(0)>0 then
								'// 존재 -> 사용함
								strSql = "update db_temp.dbo.tbl_lotte_MDCateGrp Set isUsing='Y' Where groupCode='" & groupCode & "' and MDCode='" & MDCode & "'"
								dbget.Execute(strSql)
								actCnt = actCnt+1
							else
								'// 없음 -> 신규등록
								strSql = "Insert into db_temp.dbo.tbl_lotte_MDCateGrp (groupCode, MDCode, SuperGroupName, GroupName) values " &_
										" ('" & groupCode & "'" &_
										", '" & MDCode & "'" &_
										", '" & html2db(SuperGroupName) & "'" &_
										", '" & html2db(GroupName) & "')"
								dbget.Execute(strSql)
								actCnt = actCnt+1
							end if
		
							rsget.Close
						Next
					end if
				next
				Set SGrpInfo = Nothing

				'##### DB 저장 처리 #####
			    If Err.Number = 0 Then
			    	dbget.CommitTrans				'커밋(정상)
			    	Response.Write "<script language=javascript>alert('" & actCnt & "건이 정상적으로 갱신되었습니다.');parent.history.go(0);</script>"
			    Else
			        dbget.RollBackTrans				'롤백(에러발생시)
			        Response.Write "<script language=javascript>alert('자료 저장 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
			    End If
			else
				Response.Write "<script language=javascript>alert('" & MDCode & " MD에 지정되어있는 MD상품군이 없습니다.\n롯데닷컴 담당자에게 문의해주세요.');</script>"
				Response.End
			end if
		'on Error Goto 0

		Set xmlDOM = Nothing
	else
		Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
		Response.End
	end if
	Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->