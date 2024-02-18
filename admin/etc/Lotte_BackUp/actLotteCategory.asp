<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/lotte/inc_dailyAuthCheck.asp" -->
<%
	'// 변수선언
	dim DispNo, DispNm, DispLrgNm, DispMidNm, DispSmlNm, DispThnNm
	dim CateCnt, CateInfo
	dim strSql, actCnt, disp_tp_cd, arrMDGrNo, lp

	actCnt = 0			'실갱신건수
	disp_tp_cd = requestCheckVar(request("disptpcd"),10)  ''"10"	'전시타입코드(10:일반매장, 11:브랜드매장, 12:전문매장)

	'// MD상품군 코드 접수
	strSql = "Select Distinct groupCode From db_temp.dbo.tbl_lotte_MDCateGrp " ''Where isUsing='Y'"
	rsget.Open strSql,dbget,1
		if Not(rsget.EOF or rsget.BOF) then
			reDim arrMDGrNo(rsget.recordCount)
			For lp=0 to (rsget.recordCount-1)
				arrMDGrNo(lp)=rsget(0)
				rsget.MoveNext
			Next
		else
			Response.Write "<script language=javascript>alert('등록된 MD상품군이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
			rsget.Close: dbget.Close: Response.End
		end if
	rsget.Close

	on Error Resume Next

	'// 트랜젝션 시작
	dbget.beginTrans

	'모든 MD사용여부 변경
	strSql = "update db_temp.dbo.tbl_lotte_Category Set isUsing='N', lastupdate=getdate() Where isUsing='Y' and disptpcd='"&disp_tp_cd&"'"
	dbget.Execute(strSql)

	'// 롯데닷컴 전시카테고리 조회
	for lp=0 to ubound(arrMDGrNo)-1
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", lotteAPIURL & "/openapi/searchDispCatListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&disp_tp_cd=" & disp_tp_cd & "&md_gsgr_no=" & arrMDGrNo(lp), false
'		rw lotteAPIURL & "/openapi/searchDispCatListOpenApi.lotte?subscriptionId=" & lotteAuthNo & "&disp_tp_cd=" & disp_tp_cd & "&md_gsgr_no=" & arrMDGrNo(lp)
'		response.end
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
			
			CateCnt = xmlDOM.getElementsByTagName("CategoryCount").item(0).text		'결과수
			if Err<>0 then
				Response.Write "<script language=javascript>alert('롯데닷컴 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
				dbget.RollBackTrans: dbget.Close: Response.End
			end if

			if CateCnt>0 then

				'// CateInfo Loop
				Set CateInfo = xmlDOM.getElementsByTagName("CategoryInfo")
				for each SubNodes in CateInfo
					DispNo		= Trim(SubNodes.getElementsByTagName("DispNo").item(0).text)		'카테고리 코드
					DispNm		= Trim(SubNodes.getElementsByTagName("DispNm").item(0).text)		'카테고리명(세세분류)
					DispLrgNm	= Trim(SubNodes.getElementsByTagName("DispLrgNm").item(0).text)		'대분류명
					DispMidNm	= Trim(SubNodes.getElementsByTagName("DispMidNm").item(0).text)		'중분류명
					DispSmlNm	= Trim(SubNodes.getElementsByTagName("DispSmlNm").item(0).text)		'소분류명
					DispThnNm	= Trim(SubNodes.getElementsByTagName("DispThnNm").item(0).text)		'세분류명

					'MD존재여부 확인
					strSql = "Select count(DispNo) From db_temp.dbo.tbl_lotte_Category Where DispNo='" & DispNo & "'"
					rsget.Open strSql,dbget,1

					if rsget(0)>0 then
						'// 존재 -> 사용함
						strSql = "update db_temp.dbo.tbl_lotte_Category "
						strSql = strSql & " Set isUsing='Y'"
						strSql = strSql & " , groupCode='" & arrMDGrNo(lp) & "'"
						strSql = strSql & " , disptpcd='"&disp_tp_cd&"'"
						strSql = strSql & " , DispNm='"&DispNm&"'"
						strSql = strSql & " , DispLrgNm='"&html2db(DispLrgNm)&"'"
						strSql = strSql & " , DispMidNm='"&html2db(DispMidNm)&"'"
						strSql = strSql & " , DispSmlNm='"&html2db(DispSmlNm)&"'"
						strSql = strSql & " , DispThnNm='"&html2db(DispThnNm)&"'"
						strSql = strSql & "  Where DispNo='" & DispNo & "'"
						dbget.Execute(strSql)
						actCnt = actCnt+1
					else
						'// 없음 -> 신규등록
						strSql = "Insert into db_temp.dbo.tbl_lotte_Category (DispNo, DispNm, DispLrgNm, DispMidNm, DispSmlNm, DispThnNm, disptpcd, groupCode) values " &_
								" ('" & DispNo & "'" &_
								", '" & html2db(DispNm) & "'" &_
								", '" & html2db(DispLrgNm) & "'" &_
								", '" & html2db(DispMidNm) & "'" &_
								", '" & html2db(DispSmlNm) & "'" &_
								", '" & html2db(DispThnNm) & "'" &_
								", '" & html2db(disp_tp_cd) & "'" &_
								", '" & arrMDGrNo(lp) & "')"
						dbget.Execute(strSql)
						actCnt = actCnt+1
					end if

					rsget.Close
				Next
				Set CateInfo = Nothing

			end if
	
			Set xmlDOM = Nothing
		else
			Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
			dbget.RollBackTrans: dbget.Close: Response.End
		end if
		Set objXML = Nothing

	Next

	'##### DB 저장 처리 #####
    If Err.Number = 0 Then
    	dbget.CommitTrans				'커밋(정상)
    	Response.Write "<script language=javascript>alert('" & actCnt & "건이 정상적으로 갱신되었습니다.');parent.history.go(0);</script>"
    Else
        dbget.RollBackTrans				'롤백(에러발생시)
        Response.Write "<script language=javascript>alert('자료 저장 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
    End If

	on Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->