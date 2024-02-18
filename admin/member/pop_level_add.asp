<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/LevelCls.asp" -->
<%
	dim level_sn, level_no, strTitle
	dim level_name, level_isDel
	level_sn = Request("level_sn")

	'// ��� ����
	dim oLevel
	set oLevel = new CLevel

	oLevel.FRectlevel_sn = level_sn

	if level_sn<>"" then
		'��� ��ȣ�� ������ ��������
		oLevel.GetLevel
		strTitle = "������� ����"
		if oLevel.FResultCount>0 then
			level_no = oLevel.FitemList(1).Flevel_no
			level_name = oLevel.FitemList(1).Flevel_name
			level_isDel = oLevel.FitemList(1).Flevel_isDel
		end if
	else
		strTitle = "������� ���"
	end if
%>
<html>
<head>
<title><%=strTitle%></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript">
<!--
	// ���� Ȯ��
	function chk_form(form)
	{
		if(!form.strLevel.value)
		{
			alert("��޹�ȣ�� �Է����ֽʽÿ�.");
			form.strLevel.focus();
			return false;
		}
		if(!form.level_no.value)
		{
			alert("�ߺ�Ȯ���� �������ֽʽÿ�.");
			return false;
		}
		if(form.level_name.value.length<1)
		{
			alert("��޸��� �Է����ֽʽÿ�.");
			form.level_name.focus();
			return false;
		}
		else
		{
			form.action="level_process.asp";
			form.target="_self";
			<% if level_sn<>"" then %>
			form.mode.value = "modi";
			<% else %>
			form.mode.value = "add";
			<% end if %>
			return true;
		}
	}

	// ����,����
	function chk_modi(md)
	{
		var ms, form = document.frm_level;

		if(md=='Y') ms="����";
		else ms="����";

		if(confirm("[<%=level_name%>]����� " + ms + "�Ͻðڽ��ϱ�?"))
		{
			form.action="level_process.asp";
			form.target="_self";
			form.mode.value = "del";
			form.level_isDel.value = md;
			form.submit();
		}
	}

	// �ߺ��˻�
	function chk_dupLevel()
	{
		var form = document.frm_level;
		if(!form.strLevel.value)
		{
			alert("��޹�ȣ�� �Է����ֽʽÿ�.");
			form.strLevel.focus();
			return false;
		}
		else
		{
			form.action="level_process.asp";
			form.target="FrameCKP";
			form.mode.value = "dp_chk";
			form.submit();
		}
	}
//-->
</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frm_level" method="POST" action="" onsubmit="return chk_form(this)" target="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="Level_sn" value="<%=level_sn%>">
<input type="hidden" name="level_isDel" value="<%=level_isDel%>">
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
	<td bgcolor="F8F8F8" align="center">��޹�ȣ</td>
	<td>
		<input type="text" name="strLevel" size="3" value="<%=level_no%>" onkeypress="document.frm_level.level_no.value=''">
		<input type="hidden" name="level_no" value="<%=level_no%>">
		<input type="button" value="�ߺ�Ȯ��" onClick="chk_dupLevel()">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">��޸�</td>
	<td>
		<input type="text" name="level_name" size="20" maxlength="60" value="<%=level_name%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" colspan="2">
		<input type="submit" value="Ȯ��">
		<%
			if level_sn<>"" then
				if level_isDel="N" then
		%>
		<input type="button" value="����" onClick="chk_modi('Y')">
		<%		else %>
		<input type="button" value="����" onClick="chk_modi('N')">
		<%
				end if
			end if
		%>
		<input type="button" value="���" onClick="self.close()">
	</td>
</tr>
</table>
<iframe name="FrameCKP" src="" frameborder="0" width="0" height="0"></iframe>
</form>
</body>
</html>
<% Set oLevel = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->