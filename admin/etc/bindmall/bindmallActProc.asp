<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/bindmall/bindmallCls.asp"-->
<!-- #include virtual="/admin/etc/bindmall/incbindmallFunction.asp"-->
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

If action = "REG" Then										'상품등록
	Call fnbindmallItemReg(itemid, action, resultMessage)
ElseIf action = "REGSTEP1" OR action = "REGSTEP2" Then		'상품등록 STEP
	Call fnbindmallItemRegStep(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then							'상태변경
	Call fnbindmallSellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "CHKSTAT" Then								'상품조회 + 옵션조회
	Call fnbindmallChkstat(itemid, action, resultMessage)
ElseIf action = "CHKITEM" Then								'상품조회
	Call fnbindmallChkItem(itemid, action, resultMessage)
ElseIf action = "CHKOPT" Then								'옵션조회
	Call fnbindmallChkOpt(itemid, action, resultMessage)
ElseIf action = "EDIT" Then									'상품수정
	Call fnbindmallItemEdit(itemid, action, resultMessage)
ElseIf action = "PRICE" Then								'가격수정
	Call fnbindmallItemEditPrice(itemid, action, resultMessage)
ElseIf action = "CONTENT" Then								'컨텐츠 수정
	Call fnbindmallItemEditContent(itemid, action, resultMessage)
ElseIf action = "IMAGE" Then								'이미지 수정
	Call fnbindmallItemEditIMAGE(itemid, action, resultMessage)
ElseIf action = "DELIVERY" Then								'출고/교환/반품지 수정
	Call fnbindmallItemEditDelivery(itemid, action, resultMessage)
ElseIf action = "OPTEDIT" Then								'옵션 수정
	Call fnbindmallItemEditOption(itemid, action, resultMessage)
ElseIf action = "OPTADD" Then								'옵션 수정
	Call fnbindmallItemAddOption(itemid, action, resultMessage)
ElseIf action = "updateSendState" Then						'주문상태변경 / bindmall_SongjangProc.asp에서 넘어온다.
	AssignedRow = fnbindmallSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
	response.write "<script>alert('"&AssignedRow&"건 완료 처리.');window.close()</script>"
	response.end
ElseIf action = "bindmallCommonCode" Then					'공통코드 검색
	If ccd <> "" Then
		Call fnbindmallCommonCode(ccd)
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