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
	dim CommCd
	dim page, searchDiv, searchKey, searchString, isusing, param

	dim oComm, i, lp

	'// �Ķ���� ���� //
	CommCd = RequestCheckvar(request("CommCd"),10)
	page = RequestCheckvar(request("page"),10)
	searchDiv = RequestCheckvar(request("searchDiv"),16)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = RequestCheckvar(request("searchString"),128)
	isusing = RequestCheckvar(request("isusing"),2)

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey &_
			"&searchString=" & server.URLencode(searchString) & "&isusing=" & isusing	'������ ����

	'// ���� ����
	set oComm = new CComm
	oComm.FRectCommCd = CommCd

	oComm.GetCommRead

	if (oComm.FResultCount = 0) then
	    response.write "<script>alert('�������� �ʴ� �ڵ��Դϴ�.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if
%>
<script language='javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.commNm.value)
		{
			alert("�ڵ���� �Է����ֽʽÿ�.");
			frm.commNm.focus();
			return false;
		}

		// �� ����
		return true;
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCommCd.asp">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="CommCd" value="<%=CommCd%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>�����ڵ� �� ���� / ����</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�׷�</td>
	<td width="630" bgcolor="#FFFFFF"><%=oComm.FCommList(0).FgroupNm%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�����ڵ�</td>
	<td width="630" bgcolor="#FFFFFF"><b><%=oComm.FCommList(0).FcommCd%></b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�ڵ��</td>
	<td bgcolor="#FDFDFD"><input type="text" name="commNm" value="<%=db2html(oComm.FCommList(0).FcommNm)%>" size="20" maxlength="30"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">��뿩��</td>
	<td bgcolor="#FDFDFD">
		<input type="radio" name="isUsing" value="Y" <% if oComm.FCommList(0).Fisusing="���" then Response.Write "checked"%>> ��� &nbsp; &nbsp;
		<input type="radio" name="isUsing" value="N" <% if oComm.FCommList(0).Fisusing="����" then Response.Write "checked"%>> ����
	</td>
</tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='CommCd_List.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->