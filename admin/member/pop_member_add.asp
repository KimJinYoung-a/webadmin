<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/MemberCls.asp" -->
<%
	dim id, strTitle
	dim password, company_name, email, part_sn, posit_sn, level_sn, job_sn, userdiv, isUsing,empno
	id = Request("id")

	'// ��� ����
	dim oMember
	set oMember = new CMember

	oMember.FRectid = id

	if id<>"" then
		'���� ��ȣ�� ������ ��������
		oMember.GetMember
		strTitle = "�������� ����"
		if oMember.FResultCount>0 then
			empno			= oMember.FitemList(1).Fempno
			password		= oMember.FitemList(1).Fpassword
			company_name	= oMember.FitemList(1).Fusername
			email			= oMember.FitemList(1).Fusermail
			part_sn			= oMember.FitemList(1).Fpart_sn
			posit_sn			= oMember.FitemList(1).Fposit_sn
			level_sn			= oMember.FitemList(1).Flevel_sn
			job_sn			= oMember.FitemList(1).Fjob_sn
			userdiv			= oMember.FitemList(1).Fuserdiv
			isUsing			= oMember.FitemList(1).FisUsing
		end if
	else
		strTitle = "�������� ���"
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
		
			form.action="member_process.asp";
			<% if id<>"" then %>
			form.mode.value = "modi";
			<% else %>
			form.mode.value = "add";
			<% end if %>
			return true;
		
	}

	function chk_modi(md)
	{
		var ms, form = document.frm_member;

		if(md=='N') ms="����";
		else ms="����";

		if(confirm("[<%=company_name%>]������ " + ms + "�Ͻðڽ��ϱ�?"))
		{
			form.action="member_process.asp";
			form.mode.value = "del";
			form.isUsing.value = md;
			form.submit();
		}
	}
//-->
</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frm_member" method="POST" action="" onsubmit="return chk_form(this)">
<input type="hidden" name="mode" value="">
<input type="hidden" name="isUsing" value="<%=isUsing%>">
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
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">���</td>
	<td><%=empno%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">�̸�</td>
	<td><%=company_name%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">�̸���</td>
	<td><%=email%></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">���̵�</td>
	<td>
	<% if id<>"" then %>
		<b><%=id%></b>
		<input type="hidden" name="id" value="<%=id%>">
	<% else %>
		<input type="text" name="id" size="20">
	<% end if %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">�н�����</td>
	<td>
		<input type="password" name="password" size="20" maxlength="60" value="<%=password%>">
	</td>
</tr>

<% if id<>"" and ((userdiv <= 9)  or (userdiv=111) or (userdiv=112) or (userdiv=201) or (userdiv=301)) then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">�μ�</td>
	<td>
		<%=printPartOption("part_sn", part_sn)%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">����</td>
	<td>
		<%=printPositOption("posit_sn", posit_sn)%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">��å</td>
	<td>
		<%=printJobOption("job_sn", job_sn)%>
	</td>
</tr>
<%end if%>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">���</td>
	<td>
		<%=printLevelOption("level_sn", level_sn)%>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">(����)����</td>
	<td>
		<% DrawAuthBox "userdiv",userdiv %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" colspan="2">
		<input type="submit" value="Ȯ��">
		<%
			if id<>"" then
				if isUsing="Y" then
		%>
		<input type="button" value="����" onClick="chk_modi('N')">
		<%		else %>
		<input type="button" value="����" onClick="chk_modi('Y')">
		<%
				end if
			end if
		%>
		<input type="button" value="���" onClick="self.close()">
	</td>
</tr>
</table>
</form>
</body>
</html>
<% Set oMember = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->