<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/event_off/eventAppReport.asp"-->
<%
Dim buyprice, appRunUser, appRunDay
appRunUser		= requestCheckVar(request("appRunUser"),1)
buyprice		= request("buyprice")
appRunDay		= request("appRunDay")
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">
function fnSearch(){
	opener.$("#appRunUser").val($("#appRunUser").val());
	opener.$("#buyprice").val($("#buyprice").val());
	if (!$.isNumeric($("#buyprice").val())) {
		alert('���ڸ� �Է��ϼ���');
		$("#buyprice").val("");
		$("#buyprice").focus();
		return false;
	}
	if (!$.isNumeric($("#appRunDay").val())) {
		alert('���ڸ� �Է��ϼ���');
		$("#appRunDay").val("");
		$("#appRunDay").focus();
		return false;
	}
	opener.$("#appRunDay").val($("#appRunDay").val());
	alert("������ �Ϸ�Ǿ����ϴ�.");
	opener.document.frm.submit();
	self.close();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="POST">
<tr height="30" bgcolor="#FFFFFF" id="BuyTr">
	<td colspan="11">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong>���� ��ȯ �ݾ� ����</strong>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" id="BuyTr2">
	<td width="20%">���� �ݾ�</td>
	<td align="LEFT" bgcolor="#FFFFFF"><input type="text" name="buyprice" id="buyprice" value="<%=buyprice%>"> �� �̻� ����</td>
</tr>
<tr height="30" bgcolor="#FFFFFF">
	<td colspan="11">
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr>
			<td align="left">
				<strong>���� ���� ����</strong>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="20%">����/Ȱ��ȭ ����</td>
	<td align="LEFT" bgcolor="#FFFFFF">
		<Select name="appRunUser" id="appRunUser" class="select">
			<option value="0" <%= Chkiif(appRunUser="0", "selected", "") %> >App ���� ������</option>
			<option value="1" <%= Chkiif(appRunUser="1", "selected", "") %> >��ٱ��� ���� ������</option>
		</Select>
		<input type="text" size="5" name="appRunDay" id="appRunDay" value="<%= appRunDay %>"> �� �̳� ����
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="2">
		<input type="button" value="����" class="button" onclick="fnSearch();" />
		<input type="button" value="���" class="button" onclick="self.close();"/>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->