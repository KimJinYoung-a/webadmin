<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
'###############################################
' PageName : category_left_topKeyword_write.asp
' Discription : ī�װ� žŰ���� ���/����
' History : 2008.03.31 ������ ����
'         : 2008.10.27 ��ī�װ� ó�� �߰�(������)
'         : 2009.04.15 �̹��� �߰�(������)
'###############################################

	Dim idx, cdl, cdm

	idx = Request("idx")
	cdl = Request("cdl")
	cdm = Request("cdm")

	'// ���� ����
	dim oKeyword
	Set oKeyword = new CCategoryKeyWord
	oKeyword.FRectIdx = idx

	if idx<>"" then
		oKeyword.GetCaFavKeyWord
	end if
%>
<script language="javascript">
<!--
	// ������ ���� ����
	function goSubmit()
	{
		// ī�װ� ���ÿ��� �˻�
		if(!(document.frm.cdl.value)) {
			alert("ī�װ��� �������ּ���.");
			return;
		}
		// Ű���� �Է¿��� �˻�
		if(!document.frm.keyword.value) {
			alert("���� Ű���带 �Է����ּ���.");
			document.frm.keyword.focus();
			return;
		}
		// ��ũ �Է¿��� �˻�
		if(!document.frm.linkinfo.value) {
			alert("Ű���� Ŭ���� �̵��� ��ũ�� �Է����ּ���.");
			document.frm.linkinfo.focus();
			return;
		}
		// ���� �Է¿��� �˻�
		if(!document.frm.sortNo.value) {
			alert("ǥ�� ������ �Է����ּ���.\n�� ������ �����̸� �������� ������ �����ϴ�.");
			document.frm.sortNo.focus();
			return;
		}

		<% if idx="" then %>
		if(confirm("�ۼ��Ͻ� ������ ����Ͻðڽ��ϱ�?")) {
			document.frm.mode.value="add";
			document.frm.action="DoCateTopKeyword.asp";
			document.frm.submit();
		}
		<% else %>
		if(confirm("�����Ͻ� ������ �����Ͻðڽ��ϱ�?")) {
			document.frm.mode.value="modify";
			document.frm.action="DoCateTopKeyword.asp";
			document.frm.submit();
		}
		<% end if %>
	}

	// ī�װ� ����� ���
	function changecontent() {
		if(frm.cdl.value=="110") {
			frm.action="category_left_topKeyword_write.asp";
			frm.submit();
		}
	}

	function putLinkText(key) {
		switch(key) {
			case 'search':
				document.frm.linkinfo.value='/search/search_result.asp?rect=' + document.frm.keyword.value;
				break;
			case 'cate':
				document.frm.linkinfo.value='/shopping/category_list.asp?cdl=���ڵ�&cdm=���ڵ�&cds=���ڵ�';
				break;
			case 'event':
				document.frm.linkinfo.value='/event/eventmain.asp?eventid=�̺�Ʈ��ȣ';
				break;
		}
	}
//-->
</script>
<!-- �� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="DoCateTopKeyword.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if idx="" then %>
		<font color="red"><b>žŰ���� ���</b></font>
		<% else %>
		<font color="red"><b>žŰ���� ����</b></font>
		<% end if%>
	</td>
</tr>
<% if idx<>"" then %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�Ϸù�ȣ</td>
	<td align="left"><input type="text" name="idx" value="<%=idx%>" readonly size="10" class="text_ro"></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">ī�װ�</td>
	<td align="left">
	<%
		if idx<>"" then
			DrawSelectBoxCategoryLarge "cdl", oKeyword.FitemList(0).FCDL
			if oKeyword.FitemList(0).FCDL="110" then
				if cdm="" then
					DrawSelectBoxCategoryMid "cdm", oKeyword.FitemList(0).FCDL, oKeyword.FitemList(0).FCDM
				else
					DrawSelectBoxCategoryMid "cdm", oKeyword.FitemList(0).FCDL, cdm
				end if
			end if
		else
			DrawSelectBoxCategoryLarge "cdl", cdl
			if cdl="110" then
				DrawSelectBoxCategoryMid "cdm", cdl, cdm
			end if
		end if
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">Ű����</td>
	<td align="left"><input type="text" name="keyword" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).FKeyword%>" size="32" maxlength="32" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ũ</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="text" name="linkinfo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Flinkinfo%>" size="80" maxlength="128" class="text"></td>
		<tr>
		<tr>
			<td valign="top"><font color="#707080">��)</font></td>
			<td valign="top">
				<font color="#707070">
				- <span style="cursor:pointer" onClick="putLinkText('search')">�˻���� ��ũ : /search/search_result.asp?rect=<font color="darkred">�˻���</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('cate')">ī�װ� ��ũ : /shopping/category_list.asp?cdl=<font color="darkred">���ڵ�</font>&cdm=<font color="darkred">���ڵ�</font>&cds=<font color="darkred">���ڵ�</font></span><br>
				- <span style="cursor:pointer" onClick="putLinkText('event')">�̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font></span>
				</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���û�ǰ</td>
	<td align="left">
		<input type="text" name="itemid" size="10" maxlength="8" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).Fitemid%>">
		<font color="#707070">�� ��ǰ��ȣ�� �Է�, Ű���� ���� �̹������̹Ƿ� Ű����� ������ ��� ����</font>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">ǥ�ü���</td>
	<td align="left"><input type="text" name="sortNo" value="<% if idx<>"" then Response.Write oKeyword.FitemList(0).FsortNo: else Response.Write "0" %>" size="3" class="text"></td></td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="����" onClick="goSubmit()"> &nbsp;
		<input type="button" class="button" value="���" onClick="self.history.back()">
	</td>
</tr>
</form>
<!-- �� �� -->
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
