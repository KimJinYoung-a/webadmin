<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/LtiMall/inc_dailyAuthCheck.asp"-->
<%
'// 변수선언
Dim LotteGoodNo, LotteStatCd
Dim strSql, actCnt, lp, waitCnt, rjtCnt
Dim AssignedCNT, GoodsCount

actCnt = 0			'실갱신건수
waitCnt = 0
on Error Resume Next

strSql = ""
strSql = strSql & " SELECT TOP 100 itemid, LtiMallTmpGoodNo FROM db_item.dbo.tbl_ltiMall_regItem "
strSql = strSql & " WHERE LtiMallStatCd in ('10','20','51','52') "
strSql = strSql & " and dateDiff(hh,IsNULL(lastConfirmdate,'2001-01-01'),getdate()) > 5 "
strSql = strSql & " and LtiMallTmpGoodNo is Not NULL"
strSql = strSql & " ORDER BY IsNULL(lastConfirmdate,'2001-01-01') ASC, regdate ASC"
rsget.Open strSql,dbget,1
If Not(rsget.EOF or rsget.BOF) Then
	'// 롯데아이몰 전시상품번호 매핑정보
	Do Until rsget.EOF
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "GET", ltiMallAPIURL & "/openapi/getRdToPrGoodsNoApi.lotte?subscriptionId=" & ltiMallAuthNo & "&goods_req_no=" & rsget("LtiMallTmpGoodNo"), false
			objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			objXML.Send()
			If objXML.Status = "200" Then
				Set xmlDOM = Server.CreateObject("MSXML2.DomDocument.3.0")
					xmlDOM.async = False
					xmlDOM.LoadXML BinaryToText(objXML.ResponseBody, "euc-kr")

					If Err <> 0 then
						Response.Write "<script language=javascript>alert('롯데아이몰 결과 분석 중에 오류가 발생했습니다.\n나중에 다시 시도해보세요.1');</script>"
						Set xmlDOM = Nothing
						Set objXML = Nothing
						dbget.Close: Response.End
					End If

					GoodsCount 		= Trim(xmlDOM.getElementsByTagName("GoodsCount").item(0).text)			'검색수
					LotteGoodNo		= Trim(xmlDOM.getElementsByTagName("GoodsNo").item(0).text)			'전시상품번호
					LotteStatCd		= Trim(xmlDOM.getElementsByTagName("ConfStatCd").item(0).text)		'인증상태코드

					strSql =""
					strSql = strSql & " UPDATE db_item.dbo.tbl_ltiMall_regItem "
					strSql = strSql & "	SET lastConfirmdate = getdate() "
					If (LotteStatCd <> "") then
						If LotteStatCd = "30" Then
							LotteStatCd = "7"
						End If
						strSql = strSql & "	,LtiMallStatCd='" & LotteStatCd & "' "
					End If
		
					If (LotteGoodNo > "0") and (LotteGoodNo <> "") Then
						strSql = strSql & " ,LtiMallGoodNo='" & LotteGoodNo & "' "
					End If
	
					strSql = strSql & " WHERE itemid='" & rsget("itemid") & "'"
					dbget.Execute strSql, AssignedCNT
					If (LotteStatCd = "30") Then
					    actCnt = actCnt + AssignedCNT
					ElseIf (LotteStatCd = "20") Then
					    waitCnt = waitCnt + AssignedCNT
					ElseIf (LotteStatCd = "40") Then
					    rjtCnt = rjtCnt + AssignedCNT
					End If

				Set xmlDOM = Nothing
			Else
				Response.Write "<script language=javascript>alert('롯데아이몰과 통신중에 오류가 발생했습니다.\n나중에 다시 시도해보세요.2');</script>"
				dbget.Close: Response.End
			End If
		Set objXML = Nothing
		rsget.MoveNext
	Loop
End If

rsget.Close

'##### DB 저장 처리 #####
If Err.Number = 0 Then
	If actCnt > 0 or waitCnt > 0 or rjtCnt > 0 Then
	    If (IsAutoScript) then
	        rw  "OK|"&actCnt & "건이 승인." & waitCnt& "건이 승인대기." & rjtCnt & "건 반려"
	    Else
	        If (session("ssBctID") = "icommang" or session("ssBctID") = "kjy8517") Then
	            rw actCnt & "건이 승인." & waitCnt& "건이 승인대기." & rjtCnt & "건 반려"
	        Else
    		    Response.Write "<script language=javascript>alert('" & actCnt & "건이 정상적으로 갱신되었습니다.');parent.history.go(0);</script>"
    	    End If
    	End if
	Else
	    If (IsAutoScript) Then
	        rw  "OK|"&actCnt & "건이 승인." & waitCnt& "건이 승인대기."  & rjtCnt & "건 반려"
	    Else
    		Response.Write "<script language=javascript>alert('갱신할 임시등록 상품이 없습니다.');parent.history.go(0);</script>"
    	End If
	End If
Else
    If (IsAutoScript) Then
        rw "S_ERR|처리 중에 오류가 발생했습니다"
    Else
        Response.Write "<script language=javascript>alert('처리 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
    End If
End If

on Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->