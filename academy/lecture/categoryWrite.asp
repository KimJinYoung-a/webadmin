<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
	'// ���� ���� //
	dim CateDiv
	CateDiv = RequestCheckvar(request("CateDiv"),16)
%>
<script language='javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.CateDiv.value)
		{
			alert("ī�װ� ������ �������ֽʽÿ�.");
			frm.CateDiv.focus();
			return false;
		}

		if(frm.CateCd.value.length<2)
		{
			alert("�ڵ带 �Է����ֽʽÿ�.\n\n���ڵ�� 2�ڸ��Դϴ�.");
			frm.CateCd.focus();
			return false;
		}

		if(!frm.Cate_Name.value)
		{
			alert("�ڵ���� �Է����ֽʽÿ�.");
			frm.Cate_Name.focus();
			return false;
		}

		// �� ����
		return true;
	}


	// �ڵ� �⺻�� ����
	function chgDiv(cdv)
	{
		if(cdv=='CateCD2') {
			document.all.lyEngFrm.style.display="";
		} else {
			document.all.lyEngFrm.style.display="none";
		}
	}

//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCategory.asp">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>ī�װ� �ڵ� �űԵ��</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">����</td>
	<td width="630" bgcolor="#FFFFFF">
		<select name="CateDiv" onChange="chgDiv(frm_write.CateDiv.value)">
			<option value="">����</option>
			<option value="CateCD1" <% if CateDiv="CateCD1" then Response.Write "selected" %>>Ŭ����</option>
			<option value="CateCD2" <% if CateDiv="CateCD2" then Response.Write "selected" %>>���ºо�</option>
			<option value="CateCD3" <% if CateDiv="CateCD3" then Response.Write "selected" %>>��ұ���</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">ī�װ� �ڵ�</td>
	<td width="630" bgcolor="#FFFFFF">
		<input type="text" name="CateCd" size="2" maxlength="2" value="">
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�ڵ��</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_Name" size="20" maxlength="30"></td>
</tr>
<tr id="lyEngFrm" <% if CateDiv<>"CateCD2" then Response.Write "style='display:none'" %>>
	<td align="center" width="120" bgcolor="#E8E8F1">�ڵ��(����)</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_NameEng" size="30" maxlength="40"></td>
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
