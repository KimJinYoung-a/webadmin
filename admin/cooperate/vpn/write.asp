<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	Dim vIdx
	vIdx = requestCheckVar(Request("idx"),10)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" type="text/css" href="/css/adminPartnerDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminPartnerCommon.css" />
<!--[if IE]>
	<link rel="stylesheet" type="text/css" href="/css/adminPartnerIe.css" />
<![endif]-->
<link rel="stylesheet" href="/css/scm.css" type="text/css" />
<% if session("sslgnMethod")<>"S" then %>
<!-- USBŰ ó�� ���� (2008.06.23;������) -->
<OBJECT ID='MaGerAuth' WIDTH='' HEIGHT=''	CLASSID='CLSID:781E60AE-A0AD-4A0D-A6A1-C9C060736CFC' codebase='/lib/util/MaGer/MagerAuth.cab#Version=1,0,2,4'></OBJECT>
<script language="javascript" src="/js/check_USBToken.js"></script>
<!-- USBŰ ó�� �� -->
<% end if %>
<script>
function goSaveLog(){
	if(frm1.logcont.value == ""){
		alert("�α׸� �Է��ϼ���.");
		return;
	}
	frm1.submit();
}
</script>
</head>
<body <% if session("sslgnMethod")<>"S" then %>onload="checkUSBKey()"<% end if %>>
* <strong>�α��Է�</strong>
<form name="frm1" action="proc.asp" method="post" style="margin:7px;">
<input type="hidden" name="idx" value="<%=vIdx%>">
<textarea name="logcont" cols="130" rows="14"></textarea>
<br /><br />&nbsp;&nbsp;<input type="button" value="�� ��" onclick="goSaveLog()">
</form>
</body>
</html>