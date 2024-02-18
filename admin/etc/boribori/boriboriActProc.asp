<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/boribori/boriboriCls.asp"-->
<!-- #include virtual="/admin/etc/boribori/incboriboriFunction.asp"-->
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

If action = "REG" Then										'��ǰ���
	Call fnboriboriItemReg(itemid, action, resultMessage)
ElseIf action = "REGSTEP1" OR action = "REGSTEP2" OR action = "REGSTEP3" Then	'��ǰ��� STEP
	Call fnboriboriItemRegStep(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then							'���º���
	Call fnboriboriSellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "CHKSTAT" Then								'��ǰ��ȸ
	Call fnboriboriChkstat(itemid, action, resultMessage)
ElseIf action = "EDIT" Then									'��ǰ����
	Call fnboriboriItemEdit(itemid, action, resultMessage)
ElseIf action = "PRICE" Then								'���ݼ���
	Call fnboriboriItemEditPrice(itemid, action, resultMessage)
ElseIf action = "CONTENT" Then								'�̹���&���� ����
	Call fnboriboriItemEditContent(itemid, action, resultMessage)
ElseIf action = "updateSendState" Then						'�ֹ����º��� / boribori_SongjangProc.asp���� �Ѿ�´�.
	AssignedRow = fnboriboriSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
ElseIf action = "CommCdBrand" Then							'�귣�帮��Ʈ �˻�
	Call fnboriboriBrand()
ElseIf action = "boriboriCommonCode" Then					'�����ڵ� �˻�
	If ccd <> "" Then
		Call fnBoriboriCommonCode(ccd)
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