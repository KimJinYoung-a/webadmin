<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
Dim SGrpCnt, GrpCnt, lp, SGrpInfo, GrpInfo
Dim MDCode, groupCode, SuperGroupName, GroupName
Dim strSql, actCnt
actCnt = 0		'실갱신건수
MDCode = Request("mdcd")

'// 롯데아이몰 MD상품군 조회
Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objXML.Open "GET", ltiMallAPIURL & "/openapi/searchMDGsgrListOpenApi.lotte?subscriptionId=" & ltiMallAuthNo & "&md_id=" & MDCode, false
	objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	objXML.Send()

If objXML.Status = "200" Then
	'//전달받은 내용 확인
'	Response.contentType = "text/xml; charset=euc-kr"
'	response.write BinaryToText(objXML.ResponseBody, "euc-kr")
'	response.End

	'XML을 담을 DOM 객체 생성
	Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
		xmlDOM.async = False
		xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
	'on Error Resume Next
		SGrpCnt = xmlDOM.getElementsByTagName("SuperGroupCount").item(0).text		'상위상품군 카운트
		If Err <> 0 then
			Response.Write "<script language=javascript>alert('롯데아이몰 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
			Response.End
		end if

		If SGrpCnt > 0 Then
			dbget.beginTrans
			strSql = "UPDATE db_temp.dbo.tbl_lotteiMall_MDCateGrp SET isUsing = 'N', lastupdate = getdate() WHERE isUsing = 'Y' and MDCode = '"&MDCode&"'"
			dbget.Execute(strSql)

			Set SGrpInfo = xmlDOM.getElementsByTagName("SuperGroupInfo")
			For each SGNodes in SGrpInfo
				SuperGroupName	= Trim(SGNodes.getElementsByTagName("SuperGroupName").item(0).text)		'상위상품군명
				GrpCnt			= SGNodes.getElementsByTagName("SubGroupCount").item(0).text			'하위상품군수
				If GrpCnt > 0 Then
					Set GrpInfo = SGNodes.getElementsByTagName("SubGroupInfo")
					For each SubNodes in GrpInfo
						groupCode	= Trim(SubNodes.getElementsByTagName("GroupCode").item(0).text)		'그룹코드
						GroupName	= Trim(SubNodes.getElementsByTagName("GroupName").item(0).text)		'그룹명
						strSql = "SELECT count(*) FROM db_temp.dbo.tbl_lotteiMall_MDCateGrp WHERE groupCode = '" & groupCode & "' and MDCode='" & MDCode & "'"
						rsget.Open strSql,dbget,1
						If rsget(0) > 0 Then
							strSql = "UPDATE db_temp.dbo.tbl_lotteiMall_MDCateGrp SET isUsing='Y' WHERE groupCode = '" & groupCode & "' and MDCode='" & MDCode & "'"
							dbget.Execute(strSql)
							actCnt = actCnt+1
						Else
							strSql = "INSERT INTO db_temp.dbo.tbl_lotteiMall_MDCateGrp (groupCode, MDCode, SuperGroupName, GroupName) VALUES " &_
									" ('" & groupCode & "'" &_
									", '" & MDCode & "'" &_
									", '" & html2db(SuperGroupName) & "'" &_
									", '" & html2db(GroupName) & "')"
							dbget.Execute(strSql)
							actCnt = actCnt+1
						End If
						rsget.Close
					Next
				End If
			Next
			Set SGrpInfo = Nothing

			'##### DB 저장 처리 #####
		    If Err.Number = 0 Then
		    	dbget.CommitTrans				'커밋(정상)
		    	Response.Write "<script language=javascript>alert('" & actCnt & "건이 정상적으로 갱신되었습니다.');parent.history.go(0);</script>"
		    Else
		        dbget.RollBackTrans				'롤백(에러발생시)
		        Response.Write "<script language=javascript>alert('자료 저장 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
		    End If
		Else
			Response.Write "<script language=javascript>alert('" & MDCode & " MD에 지정되어있는 MD상품군이 없습니다.\n롯데닷컴 담당자에게 문의해주세요.');</script>"
			Response.End
		End If
	'on Error Goto 0
	Set xmlDOM = Nothing
Else
	Response.Write "<script language=javascript>alert('롯데아이몰과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
	Response.End
End If
Set objXML = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->