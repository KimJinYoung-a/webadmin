<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/classes/mardy_Scrapcls.asp"-->
<%
	'// ���� ���� //
	dim ScrapId
	dim page, searchKey, searchString, param
	dim oScrap, i, lp

	'// �Ķ���� ���� //
	ScrapId = RequestCheckvar(request("ScrapId"),10)
	page = RequestCheckvar(request("page"),10)
	searchKey = RequestCheckvar(request("searchKey"),16)
	searchString = request("searchString")
  	if searchString <> "" then
		if checkNotValidHTML(searchString) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');history.back();"
		response.write "</script>"
		response.End
		end if
	end if

	'// ���� ���� ����
	set oScrap = new CMardyScrap
	oScrap.FRectScrapId = ScrapId

	oScrap.GetMardyScrapView
%>
<script language='javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.subCont.value)
		{
			alert("������ �Է����ֽʽÿ�.");
			frm.subCont.focus();
			return false;
		}

		// �� ����
		return true;
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" colspan="2" align="left">���� ��ũ�� �⺻ ����</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF"><%=oScrap.FItemList(0).Ftitle%></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF">��� ����</td>
	<td bgcolor="#FFFFFF">Type <%=oScrap.FItemList(0).FprintType%></td>
</tr>
</table>
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="http://image.thefingers.co.kr/linkweb/doMardyScrapSub.asp" enctype="multipart/form-data">
<input type="hidden" name="mode" value="wirte_sub">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="ScrapId" value="<%=ScrapId%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<input type="hidden" name="printType" value="<%=oScrap.FItemList(0).FprintType%>">
<input type="hidden" name="adminId" value="<%=Session("ssBctId")%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>���� ��ũ�� ����¹� �߰�</b></td>
</tr>
<%
	'���º� �б�
	Select Case oScrap.FItemList(0).FprintType
		Case "A"
			'///// Type A /////
%>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">����</td>
		<td bgcolor="#FFFFFF"><input type="text" name="subName" size="80" maxlength="120"></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
		<td bgcolor="#FFFFFF"><textarea name="subCont" rows="4" cols="80"></textarea></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">�̹��� #1</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="imgFile1" size="60"><br>
			<font color=darkred>�� 4:3(����:����) ������ JPG/GIF �����Դϴ�.</font>
		</td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">�̹��� #2</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="imgFile2" size="60"><br>
			<font color=darkred>�� 4:3(����:����) ������ JPG/GIF �����Դϴ�.</font>
		</td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">�̹��� #3</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="imgFile3" size="60"><br>
			<font color=darkred>�� 4:3(����:����) ������ JPG/GIF �����Դϴ�.</font>
		</td>
	</tr>
<%

		Case "B"
			'///// Type B /////
%>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
		<td bgcolor="#FFFFFF"><textarea name="subCont" rows="4" cols="80"></textarea></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">�̹��� #1</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="imgFile1" size="60"><br>
			<font color=darkred>�� 1:1(���簢��) ������ JPG/GIF �����Դϴ�.</font>
		</td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">�̹��� #2</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="imgFile2" size="60"><br>
			<font color=darkred>�� 1:1(���簢��) ������ JPG/GIF �����Դϴ�.</font>
		</td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">�̹��� #3</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="imgFile3" size="60"><br>
			<font color=darkred>�� 1:1(���簢��) ������ JPG/GIF �����Դϴ�.</font>
		</td>
	</tr>
<%
		Case "C"
			'///// Type C /////
%>

	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
		<td bgcolor="#FFFFFF"><textarea name="subCont" rows="5" cols="80"></textarea></td>
	</tr>
	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF">�̹���</td>
		<td bgcolor="#FFFFFF">
			<input type="file" name="imgFile1" size="60"><br>
			<font color=darkred>�� 17:12(����:����) ������ JPG/GIF �����Դϴ�.</font>
		</td>
	</tr>
<%
		Case "D"
			'///// Type D /////
%>

	<tr>
		<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
		<td bgcolor="#FFFFFF"><textarea name="subCont" rows="12" cols="80"></textarea></td>
	</tr>
<%
	End Select
%>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_confirm.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->