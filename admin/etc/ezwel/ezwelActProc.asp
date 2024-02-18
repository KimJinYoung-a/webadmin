<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/ezwel/ezwelCls.asp"-->
<!-- #include virtual="/admin/etc/ezwel/incEzwelFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, failCnt, chgSellYn, goodsGrpCd
Dim resultMessage, strSql, AssignedRow
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
goodsGrpCd		= request("goodsGrpCd")
failCnt			= 0

If action = "REG" Then									'상품등록
	Call fnEzwelItemReg(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then						'상태변경
	Call fnEzwelSellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "CHKSTAT" Then							'상품조회
	Call fnEzwelChkstat(itemid, action, resultMessage)
ElseIf action = "EDIT" Then								'상품수정
	Call fnEzwelItemEdit(itemid, action, resultMessage)
ElseIf action = "layout" Then							'정보고시 조회
	strSql = "DELETE FROM db_etcmall.[dbo].[tbl_ezwel_layoutList] " & VbCrlf
	strSql = strSql& " WHERE goodsGrpCd='" & goodsGrpCd & "'" & VbCrlf
	dbget.execute(strSql)

	Call fnEzwelLayout(goodsGrpCd, action, resultMessage)
	response.end
ElseIf action = "updateSendState" Then					'주문상태변경 / Ezwel_SongjangProc.asp에서 넘어온다.
	AssignedRow = fnEzwelSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
End If

response.write  "<script>" & vbCrLf &_
				"	var str, t; " & vbCrLf &_
				"	t = parent.document.getElementById('actStr') " & vbCrLf &_
				"	str = t.innerHTML; " & vbCrLf &_
				"	str += '"&resultMessage&"<br>' " & vbCrLf &_
				"	t.innerHTML = str; " & vbCrLf &_
				"	setTimeout('parent.loadRotation()', 200);" & vbCrLf &_
				"</script>"
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->