<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/PositInfoCls.asp" -->
<%
	dim posit_sn, strTitle
	dim posit_name, posit_isDel
	posit_sn = Request("posit_sn")

	'// 직급 정보
	dim oPosit
	set oPosit = new CPosit

	oPosit.FRectposit_sn = posit_sn

	if posit_sn<>"" then
		'직급 번호가 있을때 내용접수
		oPosit.GetPositInfo
		strTitle = "직급정보 수정"
		if oPosit.FResultCount>0 then
			posit_name = oPosit.FitemList(1).Fposit_name
			posit_isDel = oPosit.FitemList(1).Fposit_isDel
		end if
	else
		strTitle = "직급정보 등록"
	end if
%>
<html>
<head>
<title><%=strTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript">
<!--
	function chk_form(form)
	{
		if(form.posit_name.value.length<1)
		{
			alert("직급명을 입력해주십시요.");
			form.posit_name.focus();
			return false;
		}
		else
		{
			form.action="positInfo_process.asp";
			<% if posit_sn<>"" then %>
			form.mode.value = "modi";
			<% else %>
			form.mode.value = "add";
			<% end if %>
			return true;
		}
	}

	function chk_modi(md)
	{
		var ms, form = document.frm_positInfo;

		if(md=='Y') ms="삭제";
		else ms="복구";

		if(confirm("[<%=posit_name%>]직급을 " + ms + "하시겠습니까?"))
		{
			form.action="positInfo_process.asp";
			form.mode.value = "del";
			form.posit_isDel.value = md;
			form.submit();
		}
	}
//-->
</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frm_positInfo" method="POST" action="" onsubmit="return chk_form(this)">
<input type="hidden" name="mode" value="">
<input type="hidden" name="posit_isDel" value="<%=posit_isDel%>">
<table width="350" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr height="10" valign="bottom" bgcolor="F4F4F4">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="bottom" bgcolor="F4F4F4">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" bgcolor="F4F4F4"><b><%=strTitle%></b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="350" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<% if posit_sn<>"" then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">직급번호</td>
	<td>
		<b><%=posit_sn%></b>
		<input type="hidden" name="posit_sn" value="<%=posit_sn%>">
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">직급명</td>
	<td>
		<input type="text" name="posit_name" size="20" maxlength="60" value="<%=posit_name%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" colspan="2">
		<input type="submit" value="확인">
		<%
			if posit_sn<>"" then
				if posit_isDel="N" then
		%>
		<input type="button" value="삭제" onClick="chk_modi('Y')">
		<%		else %>
		<input type="button" value="복구" onClick="chk_modi('N')">
		<%
				end if
			end if
		%>
		<input type="button" value="취소" onClick="self.close()">
	</td>
</tr>
</table><br>
</form>
</body>
</html>
<% Set oPosit = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->