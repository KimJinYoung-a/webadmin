<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/CommCd_cls.asp"-->
<%
	'// ���� ���� //
	dim oComm, i, lp, searchDiv

	searchDiv = RequestCheckvar(request("searchDiv"),16)

	'// Ŭ���� ����
	set oComm = new CComm
%>
<script language='javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.groupCd.value)
		{
			alert("�׷��� �������ֽʽÿ�.");
			frm.groupCd.focus();
			return false;
		}

		if(frm.commCd.value.length<4)
		{
			alert("�����ڵ带 �Է����ֽʽÿ�.\n\n���ڵ�� 4�ڸ��Դϴ�.");
			frm.commCd.focus();
			return false;
		}

		if(!frm.commNm.value)
		{
			alert("�ڵ���� �Է����ֽʽÿ�.");
			frm.commNm.focus();
			return false;
		}

		// �� ����
		return true;
	}


	// �ڵ� �⺻�� ����
	function chgGrpCd(gcd)
	{
		document.frm_write.commCd.value= gcd.substring(0,1);
	}


	// �ڵ� �ߺ� �˻�
	function chkDuple(ccd)
	{
		if(ccd.length<4)
		{
			alert("�����ڵ带 �Է����ֽʽÿ�.\n\n���ڵ�� 4�ڸ��Դϴ�.");
			return;
		}
		else
		{
			FrameCHK.location = "inc_chk_commCd.asp?commCd=" + ccd;
		}
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCommCd.asp">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>�����ڵ� �űԵ��</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�׷�</td>
	<td width="630" bgcolor="#FFFFFF">
		<select name="groupCd" onChange="chgGrpCd(frm_write.groupCd.value)">
			<option value="">��ü</option>
			<%= oComm.optGroupCd(searchDiv)%>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�����ڵ�</td>
	<td width="630" bgcolor="#FFFFFF">
		<input type="text" name="commCd" size="4" maxlength="4" value="<%=left(searchDiv,1)%>">
		<img src="/images/icon_1.gif" width="55" height="21" border="0" onClick="chkDuple(frm_write.commCd.value)" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�ڵ��</td>
	<td bgcolor="#FDFDFD"><input type="text" name="commNm" size="20" maxlength="30"></td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<iframe name="FrameCHK" src="" frameborder="0" width="0" height="0"></iframe>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->
