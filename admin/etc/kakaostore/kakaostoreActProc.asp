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

If action = "IMAGE" Then								'�̹������
	Call fnkakaostoreItemImageReg(itemid, action, resultMessage)
ElseIf action = "REG" Then								'��ǰ���
	Call fnkakaostoreItemReg(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then						'���º���
	Call fnkakaostoreSellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "CHKSTAT" Then							'��ǰ��ȸ
	Call fnkakaostoreChkstat(itemid, action, resultMessage)
ElseIf action = "EDIT" Then								'��ǰ����
	Call fnkakaostoreItemEdit(itemid, action, resultMessage)
ElseIf action = "updateSendState" Then					'�ֹ����º��� / kakaostore_SongjangProc.asp���� �Ѿ�´�.
	AssignedRow = fnkakaostoreSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
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