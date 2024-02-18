<%@ language=vbscript %>
<% option explicit %>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/giftManager/GiftManagerCls.asp"-->
<% 

dim Frmvalue,imageName

Frmvalue= request("Frmvalue")
imageName = request("imageName")


%>
<%'= Frmvalue %><%'= imageName %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script language="JavaScript" src="/js/xl.js"></script>
<script language="JavaScript" src="/js/common.js"></script>
<script language="JavaScript" src="/js/report.js"></script>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
</head>
<body leftmargin="0">

<script language="javascript">
function subchk(){
	if (isNaN(UpdateFRM.viewidx.value)) {
		alert('숫자만 입력가능합니다');
		return false;
	}
}
function popInputImg(){
	var pop = window.open('','','');
}
</script>
<table width="290" border="0" cellpadding="2" cellspacing="1" class="a"  align="center" bgcolor="<%= adminColor("tablebg") %>">
	<form name="upimgfrm" method="post" action="<%= uploadUrl %>/linkweb/SpecialGift_image_Process.asp" enctype="MULTIPART/FORM-DATA">
	<input type="hidden" name="Frmvalue" value="<%= Frmvalue %>">
	<input type="hidden" name="imageName" value="<%= imageName %>">
	<tr>
		<td bgcolor="#FFFFFF" align="center"><input type="file" name="imagefile" class="button" size="26"></td>
	</tr>
	<tr>
		<td bgcolor="#FFFFFF" align="center"><input type="submit" class="button" value="저장"></td>
	</tr>
	</form>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->