<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/categoryCls.asp"-->
<%
	'// ���� ���� //
	dim CateCd
	dim page, searchKey, searchString, isusing, param, CateDiv , code_large
	dim oCate, i, lp

	'// �Ķ���� ���� //
	CateDiv = RequestCheckvar(request("CateDiv"),16)
	CateCd = RequestCheckvar(request("CateCd"),3)
	code_large = RequestCheckvar(request("code_large"),3)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
	isusing = RequestCheckvar(request("isusing"),1)
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if
	param = "&page=" & page & "&searchKey=" & searchKey & "&CateDiv=" & CateDiv &_
			"&searchString=" & server.URLencode(searchString) & "&isusing=" & isusing	'������ ����

	'// ���� ����
	set oCate = new CCate
	oCate.FCateDiv = CateDiv
	oCate.FRectCateCd = CateCd
	oCate.FRectLargeCateCd = code_large
	oCate.GetCateRead

	if (oCate.FResultCount = 0) then
	    response.write "<script>alert('�������� �ʴ� �ڵ��Դϴ�.'); history.back();</script>"
	    dbget.close()	:	response.End
	end if
%>
<script language='javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.Cate_Name.value)
		{
			alert("�ڵ���� �Է����ֽʽÿ�.");
			frm.Cate_Name.focus();
			return false;
		}

		// �� ����
		return true;
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="750" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doCategory2012.asp">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="CateDiv" value="<%=CateDiv%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="code_large" value="<%=code_large%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>ī�װ� �ڵ� �� ���� / ����</b></td>
</tr>
<% if CateDiv="code_mid" then %>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">��ī�װ� �ڵ�</td>
	<td width="630" bgcolor="#FFFFFF"><b><%=oCate.FCateList(0).FcateLargeCd%></b></td>
</tr>
<% End If %>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">ī�װ� �ڵ�</td>
	<td width="630" bgcolor="#FFFFFF"><b><input type="text" name="CateCd" value="<%=oCate.FCateList(0).FCateCd%>" size="2"></b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�ڵ��</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_Name" value="<%=db2html(oCate.FCateList(0).FCateCD_Name)%>" size="20" maxlength="30"></td>
</tr>
<% if CateDiv="code_mid" then %>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">�ڵ��(����)</td>
	<td bgcolor="#FDFDFD"><input type="text" name="Cate_NameEng" value="<%=db2html(oCate.FCateList(0).FCateCD_NameEng)%>" size="30" maxlength="40"></td>
</tr>
<% end if %>
<% if CateDiv="code_mid" then %>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">ǥ�ü���</td>
	<td bgcolor="#FDFDFD"><input type="text" name="orderno" value="<%=db2html(oCate.FCateList(0).FsortNo)%>" size="3"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#E8E8F1">��뿩��</td>
	<td bgcolor="#FDFDFD">
		<input type="radio" name="isUsing" value="Y" <% if oCate.FCateList(0).Fisusing="���" then Response.Write "checked"%>> ��� &nbsp; &nbsp;
		<input type="radio" name="isUsing" value="N" <% if oCate.FCateList(0).Fisusing="����" then Response.Write "checked"%>> ����
	</td>
</tr>
<% end if %>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="self.location='CategoryList.asp?menupos=<%=menupos & param%>'" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->