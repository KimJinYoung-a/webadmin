<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [CS]��������>>[FAQ]���� 
' Hieditor : 2009.03.02 �̿��� ����
'			 2021.07.30 �ѿ�� ����(��뿩�� �߰�)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/faq_cls.asp"-->
<%
	'// ���� ���� //
	dim faqid, isusing
	dim page, searchDiv, searchKey, searchString, param

	dim ofaq, i, lp

	'// �Ķ���� ���� //
	faqid = request("faqid")
	page = request("page")
	searchDiv = request("searchDiv")
	searchKey = request("searchKey")
	searchString = request("searchString")
	isusing = requestcheckvar(request("isusing"),1)

	if page="" then page=1
	if searchKey="" then searchKey="title"

	param = "&page=" & page & "&searchDiv=" & searchDiv & "&searchKey=" & searchKey & "&searchString=" & searchString	'������ ����

	'// ���� ����
	set ofaq = new Cfaq
	ofaq.FRectfaqid = faqid

	ofaq.GetFAQRead

isusing = ofaq.FfaqList(0).fisusing
if isusing="" or isnull(isusing) then isusing="Y"
%>
<script type='text/javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.commCd.value)
		{
			alert("������ �������ֽʽÿ�.");
			frm.commCd.focus();
			return false;
		}

		if(!frm.title.value)
		{
			alert("������ �Է����ֽʽÿ�.");
			frm.title.focus();
			return false;
		}

		if(!frm.contents.value)
		{
			alert("������ �ۼ����ֽʽÿ�.");
			frm.contents.focus();
			return false;
		}
		if(!frm.isusing.value){
			alert("��뿩�θ� ������ �ּ���.");
			frm.isusing.focus();
			return false;
		}

		// �� ����
		return true;
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="faq_process.asp">
<input type="hidden" name="mode" value="UPD">
<input type="hidden" name="faqId" value="<%=faqId%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="searchDiv" value="<%=searchDiv%>">
<input type="hidden" name="searchKey" value="<%=searchKey%>">
<input type="hidden" name="searchString" value="<%=searchString%>">
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<tr align="center" height="25" bgcolor="#F0F0FD">
	<td colspan="2"><b>faq ���� ����</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="commCd">
			<option value="">����</option>
			<%= db2html(ofaq.optCommCd("Z200", ofaq.FfaqList(0).FcommCd)) %>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>���ļ���</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="disporder" value="<%= ofaq.FfaqList(0).Fdisporder %>" size="3" maxlength="3">&nbsp;&nbsp;�����Է�(0-999)���̰�</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="title" value="<%= ReplaceBracket(db2html(ofaq.FfaqList(0).Ftitle)) %>" size="80" maxlength="80"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><textarea name="contents" class="textarea" rows="14" cols="80"><%= ReplaceBracket(db2html(ofaq.FfaqList(0).Fcontents)) %></textarea></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>Link��</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="linkname" value="<%= ReplaceBracket(db2html(ofaq.FfaqList(0).Flinkname)) %>" size="30" maxlength="30"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>LinkURL</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="linkurl" value="<%= ReplaceBracket(db2html(ofaq.FfaqList(0).Flinkurl)) %>" size="80" maxlength="80"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>��뿩��</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxUsingYN "isusing", isusing %></td>
</tr>
<tr align="center" height="25" bgcolor="#F0F0FD">
	<td colspan="2">
		<input type="submit" class="button" value="�����ϱ�">
		<input type="button" class="button" value="����ϱ�" onClick="history.back()">
	</td>
</tr>
</table>
</form>
<!-- ���� ȭ�� �� -->
<%
	set ofaq = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->