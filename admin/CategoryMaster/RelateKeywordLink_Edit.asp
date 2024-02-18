<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : RelateKeywordLink_Edit.asp
' Discription : ī�װ� ���� Ű���� ���/����
' History : 2008.03.28 ������ ����
'			2022.07.05 �ѿ�� ����(isms�������ġ, ǥ���ڵ����κ���)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
	Dim LinkCode
	LinkCode = requestcheckvar(getNumeric(request("rid")),10)

	'// ���� ����
	dim oRelate
	Set oRelate = new CRelateList
	oRelate.FRectLinkCode = LinkCode

	if LinkCode<>"" then
		oRelate.GetRelateLinkItem
	end if
%>
<script type='text/javascript'>
<!--
	// ������ ���� ����
	function goSubmit()
	{
		// ī�װ� �ߺз����� �Է��ߴ��� �˻�
		if(!(document.frm.cdl.value&&document.frm.cdm.value)) {
			alert("ī�װ��� �������ּ���.\n\n�� ���� Ű����� ī�װ� �ߺз����� �����ϼž��մϴ�.");
			return;
		}
		// Ű���� �Է¿��� �˻�
		if(!document.frm.linkKeyword.value) {
			alert("���� Ű���带 �Է����ּ���.");
			document.frm.linkKeyword.focus();
			return;
		}
		// ��ũ �Է¿��� �˻�
		if(!document.frm.linkURL.value) {
			alert("Ű���� Ŭ���� �̵��� ��ũ�� �Է����ּ���.");
			document.frm.linkURL.focus();
			return;
		}

		<% if LinkCode="" then %>
		if(confirm("�ۼ��Ͻ� ������ ����Ͻðڽ��ϱ�?")) {
			document.frm.mode.value="add";
			document.frm.action="DoRelate_Process.asp";
			document.frm.submit();
		}
		<% else %>
		if(confirm("�����Ͻ� ������ �����Ͻðڽ��ϱ�?")) {
			document.frm.mode.value="modify";
			document.frm.action="DoRelate_Process.asp";
			document.frm.submit();
		}
		<% end if %>
	}
//-->
</script>
<!-- �� ���� -->
<form name="frm" method="get" action="" action="DoRelate_Process.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="30">
	<td colspan="2" bgcolor="#FFFFFF">
		<img src="/images/icon_star.gif" align="absmiddle">
		<% if LinkCode="" then %>
		<font color="red"><b>���� Ű���� ���</b></font>
		<% else %>
		<font color="red"><b>���� Ű���� ����</b></font>
		<% end if%>
	</td>
</tr>
<% if LinkCode<>"" then %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ũ�ڵ�</td>
	<td align="left"><input type="text" name="rid" value="<%=LinkCode%>" readonly size="10" class="text_ro"></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">ī�װ�</td>
	<td align="left">
		<%
			'ī�װ� ����
			if LinkCode<>"" then
				tmp_cdl = oRelate.FitemList(1).FcdL
				tmp_cdm = oRelate.FitemList(1).FcdM
				tmp_cds = oRelate.FitemList(1).FcdS
			end if
		%>
		<!-- #include virtual="/common/module/categoryselectbox.asp"-->
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">Ű����</td>
	<td align="left"><input type="text" name="linkKeyword" value="<% if LinkCode<>"" then Response.Write ReplaceBracket(oRelate.FitemList(1).FlinkKeyword) %>" size="32" maxlength="32" class="text"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��ũ</td>
	<td align="left">
		<table cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td colspan="2"><input type="text" name="linkURL" value="<% if LinkCode<>"" then Response.Write ReplaceBracket(oRelate.FitemList(1).FlinkURL) %>" size="80" maxlength="128" class="text"></td>
		<tr>
		<tr>
			<td valign="top"><font color="#707080">��)</font></td>
			<td valign="top">
				<font color="#707070">
				- ī�װ� ��ũ : /shopping/category_list.asp?cdl=<font color="darkred">���ڵ�</font>&cdm=<font color="darkred">���ڵ�</font>&cds=<font color="darkred">���ڵ�</font><br>
				- �̺�Ʈ ��ũ : /event/eventmain.asp?eventid=<font color="darkred">�̺�Ʈ�ڵ�</font>
				</font>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#FFFFFF">
		<input type="button" class="button" value="����" onClick="goSubmit()"> &nbsp;
		<input type="button" class="button" value="���" onClick="self.history.back()">
	</td>
</tr>
<!-- �� �� -->
</table>
</form>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->