<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/kakaostore/kakaostoreCls.asp"-->
<!-- #include virtual="/admin/etc/kakaostore/inckakaostoreFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, failCnt, chgSellYn, goodsGrpCd
Dim resultMessage, strSql, AssignedRow, ccd
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
goodsGrpCd		= request("goodsGrpCd")
ccd				= request("ccd")
failCnt			= 0

If action = "IMAGE" Then								'이미지등록
	Call fnkakaostoreItemImageReg(itemid, action, resultMessage)
ElseIf action = "REG" Then								'상품등록
	Call fnkakaostoreItemReg(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then						'상태변경
	Call fnkakaostoreSellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "CHKSTAT" Then							'상품조회
	Call fnkakaostoreChkstat(itemid, action, resultMessage)
ElseIf action = "EDIT" Then								'상품수정
	Call fnkakaostoreItemEdit(itemid, action, resultMessage)
ElseIf action = "updateSendState" Then					'주문상태변경 / kakaostore_SongjangProc.asp에서 넘어온다.
	AssignedRow = fnkakaostoreSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
ElseIf action = "kakaostoreCommonCode" Then
	If ccd = "category" Then
		Call fnkakaostoreCategory(ccd)
	End If
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