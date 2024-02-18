<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->
<%
	Dim mid

	mid = Request("mid")
%>
<html>
<head>
<title>메뉴 권한 선택</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript">
<!--
	function putItem()
	{
		var form=document.frm_auth;
		if(!form.part_sn.value)
		{
			alert("부서를 선택해주십시오.");
			form.part_sn.focus();
			return;
		}
		if(!form.level_sn.value)
		{
			alert("등급을 선택해주십시오.");
			form.level_sn.focus();
			return;
		}
		else
		{
			psn = form.part_sn.options[form.part_sn.selectedIndex].value;
			pnm = form.part_sn.options[form.part_sn.selectedIndex].text;
			lsn = form.level_sn.options[form.level_sn.selectedIndex].value;
			lnm = form.level_sn.options[form.level_sn.selectedIndex].text;
			opener.addAuthItem(psn,pnm,lsn,lnm);
			self.close();
		}
	}
//-->
</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frm_auth" method="GET" action="">
<table width="350" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr height="10" valign="bottom" bgcolor="F4F4F4">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="bottom" bgcolor="F4F4F4">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" bgcolor="F4F4F4"><b>메뉴 권한 선택</b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="350" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<% if mid<>"" then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">메뉴번호</td>
	<td><%=mid%></td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">부 서</td>
	<td><%=printPartOption("part_sn", "")%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">등 급</td>
	<td><%=printLevelOption("level_sn", "")%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" colspan="2">
		<input type="button" value="확인" onClick="putItem()">
		<input type="button" value="취소" onClick="self.close()">
	</td>
</tr>
</table><br>
</form>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->