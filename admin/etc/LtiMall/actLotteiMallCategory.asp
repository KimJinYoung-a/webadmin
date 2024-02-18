<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
Dim DispNo, DispNm, DispLrgNm, DispMidNm, DispSmlNm, DispThnNm
Dim CateCnt, CateInfo
Dim strSql, actCnt, disp_tp_cd, arrMDGrNo, lp
actCnt = 0			'실갱신건수
disp_tp_cd = requestCheckVar(request("disptpcd"),10)  ''"10"	'전시타입코드(10:일반매장, 11:브랜드매장, 12:전문매장)
'// MD상품군 코드 접수
strSql = "SELECT Distinct groupCode FROM db_temp.dbo.tbl_lotteiMall_MDCateGrp WHERE isUsing = 'Y'"
rsget.Open strSql,dbget,1
If Not(rsget.EOF or rsget.BOF) then
	ReDim arrMDGrNo(rsget.recordCount)
	For lp = 0 to (rsget.recordCount - 1)
		arrMDGrNo(lp)=rsget(0)
		rsget.MoveNext
	Next
Else
	Response.Write "<script language=javascript>alert('등록된 MD상품군이 없습니다.\n확인 후 다시 시도해주세요.');</script>"
	rsget.Close: dbget.Close: Response.End
End If
rsget.Close

'on Error Resume Next
dbget.beginTrans

'모든 MD사용여부 변경
strSql = "update db_temp.dbo.tbl_lotteiMall_Category Set isUsing='N', lastupdate=getdate() Where isUsing='Y' and disptpcd='"&disp_tp_cd&"'"
dbget.Execute(strSql)
'// 롯데아이몰 전시카테고리 조회
for lp = 0 to ubound(arrMDGrNo)-1
	Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
		objXML.Open "GET", ltiMallAPIURL & "/openapi/searchDispCatListOpenApi.lotte?subscriptionId=" & ltiMallAuthNo & "&disp_tp_cd=" & disp_tp_cd & "&md_gsgr_no=" & arrMDGrNo(lp), false
		objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objXML.Send()
		If objXML.Status = "200" Then
			Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
				xmlDOM.async = False
				xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")
				CateCnt = xmlDOM.getElementsByTagName("CategoryCount").item(0).text		'결과수
				If Err <> 0 Then
					Response.Write "<script language=javascript>alert('롯데아이몰 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
					dbget.RollBackTrans: dbget.Close: Response.End
				End If

				If CInt(CateCnt) > 0 Then
					Set CateInfo = xmlDOM.getElementsByTagName("CategoryInfoList")
						For each SubNodes in CateInfo
							DispNo		= Trim(SubNodes.getElementsByTagName("DispNo").item(0).text)		'카테고리 코드
							DispNm		= Trim(SubNodes.getElementsByTagName("DispNm").item(0).text)		'카테고리명(세세분류)
							DispLrgNm	= Trim(SubNodes.getElementsByTagName("DispLrgNm").item(0).text)		'대분류명
							DispMidNm	= Trim(SubNodes.getElementsByTagName("DispMidNm").item(0).text)		'중분류명
							DispSmlNm	= Trim(SubNodes.getElementsByTagName("DispSmlNm").item(0).text)		'소분류명
							DispThnNm	= Trim(SubNodes.getElementsByTagName("DispThnNm").item(0).text)		'세분류명

							'MD존재여부 확인
							strSql = "Select count(DispNo) From db_temp.dbo.tbl_lotteiMall_Category Where DispNo='" & DispNo & "' and groupCode = '" & arrMDGrNo(lp) & "' "
							rsget.Open strSql,dbget,1
							If rsget(0) > 0 Then
								'// 존재 -> 사용함
								strSql = "update db_temp.dbo.tbl_lotteiMall_Category Set isUsing='Y', groupCode='" & arrMDGrNo(lp) & "', disptpcd='"&disp_tp_cd&"' Where DispNo='" & DispNo & "'  and groupCode = '" & arrMDGrNo(lp) & "' "
								dbget.Execute(strSql)
							Else
								'// 없음 -> 신규등록
								strSql = "Insert into db_temp.dbo.tbl_lotteiMall_Category (DispNo, DispNm, DispLrgNm, DispMidNm, DispSmlNm, DispThnNm, disptpcd, groupCode) values " &_
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
							End If
							rsget.Close
						Next
					Set CateInfo = Nothing
				End If
			Set xmlDOM = Nothing
		Else
			Response.Write "<script language=javascript>alert('롯데닷컴과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요');</script>"
			dbget.RollBackTrans: dbget.Close: Response.End
		End If
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