<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 60 * 15 %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/JSON_2.0.4.asp"-->
<!-- #include virtual="/admin/etc/benepia/benepiaCls.asp"-->
<!-- #include virtual="/admin/etc/benepia/incbenepiaFunction.asp"-->
<!-- #include virtual="/admin/etc/incOutmallCommonFunction.asp"-->
<!-- #include virtual="/admin/etc/ezwel/incEzwelFunction.asp"-->
<script language="jscript" runat="server" src="/lib/util/JSON_PARSER_JS.asp"></script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#F4F4F4" >
<%
Dim strSql, xmlDOM
Dim objXML, iMessage
Dim ord_no     : ord_no     = request("ord_no")
Dim ord_dtl_sn : ord_dtl_sn = request("ord_dtl_sn")
Dim hdc_cd     : hdc_cd     = request("hdc_cd")
Dim inv_no     : inv_no     = Left(request("inv_no"), 15)					'// 15�� ������ ����
Dim failCount  : failCount     = 0
inv_no = Trim(getNumeric(inv_no))
Dim ORG_ord_no : ORG_ord_no = ord_no
dim prctp : prctp = requestCheckvar(request("prctp"),20)    ''ó�� Action (3:����Ϸ���, )

'' �ֹ��� ���� �Է��ϴ� ���̽�.
IF (InStr(ord_no, "_") > 0) then
	ord_no = getOutmallRefOrgOrderNO(ord_no, ord_dtl_sn, CMALLNAME)
End If

Call fnbenepiaSongjangUpload(ord_no, ord_dtl_sn, hdc_cd, inv_no, iMessage, failCount, ORG_ord_no)
Dim IsSuccss : IsSuccss=(iMessage="OK")
If NOT(IsSuccss) Then
    rw "<font color=red>"&iMessage&"</font>"
    rw ord_no
    rw ord_dtl_sn
    rw hdc_cd
    rw inv_no

	If failCount > 0 Then
		Dim reqURI 
		if (InStr(iMessage,"�߸��� �����ȣ �Դϴ�")>0) then
			reqURI="?ord_no="&request("ord_no")&"&ord_dtl_sn="&request("ord_dtl_sn")&"&hdc_cd=1082&inv_no="&request("inv_no")&"&isfrcsend=1"
        	response.write "<br><input type='button' value='��Ÿ��� ����' onClick=""location.href='"&reqURI&"'""><br>"
		end if
		response.write  "<select name='updateSendState' id=""updateSendState"">" &_
						"	<option value=''>����</option>" &_
						"	<option value='951'>������ ����</option>" &_
						"	<option value='952'>����ֹ�</option>" &_
						"</select>&nbsp;&nbsp;"
		
		response.write "<input type='button' value='�Ϸ�ó��' onClick=""fnSetSendState('"&ORG_ord_no&"','"&ord_dtl_sn&"',document.getElementById('updateSendState').value)"">"
		response.write "<script language='javascript'>"&VbCRLF
		response.write "function fnSetSendState(ORG_ord_no,ord_dtl_sn,selectValue){"&VbCRLF
		response.write "    if(selectValue == ''){"&VbCRLF
		response.write "    	alert('�������ּ���');"&VbCRLF
		response.write "    	document.getElementById('updateSendState').focus();"&VbCRLF
		response.write "    	return;"&VbCRLF
		response.write "    }"&VbCRLF
		response.write "    var uri = 'benepiaActProc.asp?act=updateSendState&ord_no='+ORG_ord_no+'&ord_dtl_sn='+ord_dtl_sn+'&updateSendState='+selectValue;"&VbCRLF
		response.write "    location.replace(uri);"&VbCRLF
		response.write "}"&VbCRLF
		response.write "</script>"&VbCRLF
	End If
Else
	rw "OK"
End If
%>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->