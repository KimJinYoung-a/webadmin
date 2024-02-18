<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/PartInfoCls.asp" -->
<%
	dim part_sn, strTitle
	dim part_name, part_sort, part_isDel
	part_sn = Request("part_sn")

	'// �μ� ����
	dim oPart
	set oPart = new CPart

	oPart.FRectpart_sn = part_sn

	if part_sn<>"" then
		'�μ� ��ȣ�� ������ ��������
		oPart.GetPartInfo
		strTitle = "�μ����� ����"
		if oPart.FResultCount>0 then
			part_name = oPart.FitemList(1).Fpart_name
			part_sort = oPart.FitemList(1).Fpart_sort
			part_isDel = oPart.FitemList(1).Fpart_isDel
		end if
	else
		strTitle = "�μ����� ���"
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
		if(form.part_name.value.length<1)
		{
			alert("�μ����� �Է����ֽʽÿ�.");
			form.part_name.focus();
			return false;
		}
		if(!form.part_sort.value)
		{
			alert("�μ��� ���� ������ �Է����ֽʽÿ�.");
			form.part_sort.focus();
			return false;
		}
		else
		{
			form.action="partInfo_process.asp";
			<% if part_sn<>"" then %>
			form.mode.value = "modi";
			<% else %>
			form.mode.value = "add";
			<% end if %>
			return true;
		}
	}

	function chk_modi(md)
	{
		var ms, form = document.frm_partInfo;

		if(md=='Y') ms="����";
		else ms="����";

		if(confirm("[<%=part_name%>]�μ��� " + ms + "�Ͻðڽ��ϱ�?"))
		{
			form.action="partInfo_process.asp";
			form.mode.value = "del";
			form.part_isDel.value = md;
			form.submit();
		}
	}
//-->
</script>
</head>
<body leftmargin="5" topmargin="5">
<form name="frm_partInfo" method="POST" action="" onsubmit="return chk_form(this)">
<input type="hidden" name="mode" value="">
<input type="hidden" name="part_isDel" value="<%=part_isDel%>">
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
<% if part_sn<>"" then %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">�μ���ȣ</td>
	<td>
		<b><%=part_sn%></b>
		<input type="hidden" name="part_sn" value="<%=part_sn%>">
	</td>
</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">�μ���</td>
	<td>
		<input type="text" name="part_name" size="20" maxlength="60" value="<%=part_name%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center">���Ĺ�ȣ</td>
	<td>
		<input type="text" name="part_sort" size="3" value="<%=part_sort%>">
		�� ȭ�鿡 ������ ���� ����
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" colspan="2">
		<input type="submit" value="Ȯ��">
		<%
			if part_sn<>"" then
				if part_isDel="N" then
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
</table><br>
</form>
</body>
</html>
<% Set oPart = Nothing %>
<!-- #include virtual="/lib/db/dbclose.asp" -->