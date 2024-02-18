<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/etc/wconcept/wconceptCls.asp"-->
<!-- #include virtual="/admin/etc/wconcept/incwconceptFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<%
Dim itemid, action, failCnt, chgSellYn, goodsGrpCd
Dim resultMessage, strSql, AssignedRow, ccd, arrRows, lp
itemid			= requestCheckVar(request("itemid"),9)
action			= request("act")
chgSellYn		= request("chgSellYn")
goodsGrpCd		= request("goodsGrpCd")
ccd				= request("ccd")
failCnt			= 0

If action = "REG" Then										'��ǰ���
	Call fnwconceptItemReg(itemid, action, resultMessage)
ElseIf action = "REGSTEP1" OR action = "REGSTEP2" OR action = "REGSTEP3" OR action = "REGSTEP4" OR action = "REGSTEP5" OR action = "REGSTEP6" Then		'��ǰ��� STEP
	Call fnwconceptItemRegStep(itemid, action, resultMessage)
ElseIf action = "EditSellYn" Then							'���º���
	Call fnwconceptSellYN(itemid, action, chgSellYn, resultMessage)
ElseIf action = "CHKSTAT" Then								'��ǰ��ȸ + �ɼ���ȸ
	Call fnwconceptChkstat(itemid, action, resultMessage)
ElseIf action = "CHKITEM" Then								'��ǰ��ȸ
	Call fnwconceptChkItem(itemid, action, resultMessage)
ElseIf action = "CHKOPT" Then								'�ɼ���ȸ
	Call fnwconceptChkOpt(itemid, action, resultMessage)
ElseIf action = "EDIT" Then									'��ǰ����
	Call fnwconceptItemEdit(itemid, action, resultMessage)
ElseIf action = "PRICE" Then								'���ݼ���
	Call fnwconceptItemEditPrice(itemid, action, resultMessage)
ElseIf action = "CONTENT" Then								'������ ����
	Call fnwconceptItemEditContent(itemid, action, resultMessage)
ElseIf action = "IMAGE" Then								'�̹��� ����
	Call fnwconceptItemEditIMAGE(itemid, action, resultMessage)
ElseIf action = "ADDIMAGE" Then								'���/��ȯ/��ǰ�� ����
	Call fnwconceptItemEditAddIMAGE(itemid, action, resultMessage)
ElseIf action = "OPTEDIT" Then								'�ɼ� ����
	Call fnwconceptItemEditOption(itemid, action, resultMessage)
ElseIf action = "INFOCODE" Then								'������� ����
	Call fnwconceptInfoCode(itemid, action, resultMessage)
ElseIf action = "updateSendState" Then						'�ֹ����º��� / wconcept_SongjangProc.asp���� �Ѿ�´�.
	AssignedRow = fnwconceptSongjangUploadByManager(CMALLNAME, request("ord_no"), request("ord_dtl_sn"), request("updateSendState"))
	response.write "<script>alert('"&AssignedRow&"�� �Ϸ� ó��.');window.close()</script>"
	response.end
ElseIf action = "wconceptCommonCode" Then					'�����ڵ� �˻�
	If ccd = "PRODUCT_NOTICE_INFO" Then
		strSql = ""
		strSql = strSql & " SELECT MediumCode, CategoryCode "
		strSql = strSql & " FROM db_etcmall.dbo.[tbl_wconcept_category] "
		strSql = strSql & " GROUP BY MediumCode, CategoryCode "
		strSql = strSql & " ORDER BY 1, 2 "
		rsget.CursorLocation = adUseClient
		rsget.Open strSql, dbget, adOpenForwardOnly, adLockReadOnly
		If Not(rsget.EOF or rsget.BOF) then
			arrRows = rsget.getRows()
		End If
		rsget.Close

		strSql = ""
		strSql = strSql & " DELETE FROM db_etcmall.dbo.[tbl_wconcept_infoCode] "
		dbget.Execute strSql

		For lp =0 To UBound(arrRows,2)
			goodsGrpCd = arrRows(0, lp) & "," & arrRows(1, lp)
			Call fnwconceptCommonCode(ccd, goodsGrpCd)
			response.flush
			response.clear
		Next
		rw "---END---"
	Else
		Call fnwconceptCommonCode(ccd, goodsGrpCd)
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