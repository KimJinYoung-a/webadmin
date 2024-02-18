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

If action = "REG" Then										'��ǰ���
	Call fnbindmallItemReg(itemid, action, resultMessage)
ElseIf action = "REGSTEP1" OR action = "REGSTEP2" Then		'��ǰ��� STEP
	Call fnbindmallItemRegStep(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then							'���º���
	Call fnbindmallSellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "CHKSTAT" Then								'��ǰ��ȸ + �ɼ���ȸ
	Call fnbindmallChkstat(itemid, action, resultMessage)
ElseIf action = "CHKITEM" Then								'��ǰ��ȸ
	Call fnbindmallChkItem(itemid, action, resultMessage)
ElseIf action = "CHKOPT" Then								'�ɼ���ȸ
	Call fnbindmallChkOpt(itemid, action, resultMessage)
ElseIf action = "EDIT" Then									'��ǰ����
	Call fnbindmallItemEdit(itemid, action, resultMessage)
ElseIf action = "PRICE" Then								'���ݼ���
	Call fnbindmallItemEditPrice(itemid, action, resultMessage)
ElseIf action = "CONTENT" Then								'������ ����
	Call fnbindmallItemEditContent(itemid, action, resultMessage)
ElseIf action = "IMAGE" Then								'�̹��� ����
	Call fnbindmallItemEditIMAGE(itemid, action, resultMessage)
ElseIf action = "DELIVERY" Then								'���/��ȯ/��ǰ�� ����
	Call fnbindmallItemEditDelivery(itemid, action, resultMessage)
ElseIf action = "OPTEDIT" Then								'�ɼ� ����
	Call fnbindmallItemEditOption(itemid, action, resultMessage)
ElseIf action = "OPTADD" Then								'�ɼ� ����
	Call fnbindmallItemAddOption(itemid, action, resultMessage)
ElseIf action = "updateSendState" Then						'�ֹ����º��� / bindmall_SongjangProc.asp���� �Ѿ�´�.
	AssignedRow = fnbindmallSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
ElseIf action = "bindmallCommonCode" Then					'�����ڵ� �˻�
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