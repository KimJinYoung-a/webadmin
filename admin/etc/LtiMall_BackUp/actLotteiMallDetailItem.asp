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
strSql = strSql & " SELECT TOP 100 T.ltimallgoodno, T.itemid "
strSql = strSql & " FROM db_temp.dbo.tbl_tmp_ltimallGoodno as T "
strSql = strSql & " LEFT JOIN db_item.dbo.tbl_LTiMall_regItem as L on T.ltimallgoodno = L.ltimallgoodno "
strSql = strSql & " WHERE L.itemid is NULL "
strSql = strSql & " and T.goodsRegdtime > '2013-07-01' "
strSql = strSql & " and T.itemid is null "
strSql = strSql & " ORDER BY T.goodsregdtime "
rsget.Open strSql,dbget,1
If Not(rsget.EOF or rsget.BOF) Then
	'// 롯데아이몰 전시상품번호 매핑정보
	Do Until rsget.EOF
		Set objXML= CreateObject("MSXML2.ServerXMLHTTP.3.0")
			objXML.Open "GET", ltiMallAPIURL & "/openapi/searchNewGoodsDtlInfoOpenApi.lotte?subscriptionId=" & ltiMallAuthNo & "&goods_no=" & rsget("ltimallgoodno"), false
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

					GoodsNo 		= Trim(xmlDOM.getElementsByTagName("GoodsNo").item(0).text)
					CorpGoodsNo		= Trim(xmlDOM.getElementsByTagName("CorpGoodsNo").item(0).text)

					strSql =""
					strSql = strSql & " UPDATE db_temp.dbo.tbl_tmp_ltimallGoodno "
					strSql = strSql & "	SET itemid = '"&CorpGoodsNo&"' "
					strSql = strSql & " WHERE ltimallgoodno='" & rsget("ltimallgoodno") & "'"
rw strSql
					dbget.Execute strSql, AssignedCNT

					If (CorpGoodsNo <> "") Then
					    actCnt = actCnt + AssignedCNT
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
	If actCnt > 0 Then
        If (session("ssBctID") = "icommang" or session("ssBctID") = "kjy8517") Then
            rw actCnt & "건 수정"
	    End If
	Else
		rw actCnt & "건 수정"
	End If
Else
	Response.Write "<script language=javascript>alert('처리 중에 오류가 발생했습니다.\n관리자에게 문의해주세요.');</script>"
End If

on Error Goto 0
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->