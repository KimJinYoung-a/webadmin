<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/items/itemImageTextCls.asp"-->
<%

dim itemid
itemid = requestCheckvar(request("itemid"),32)

dim oitem
set oitem = new CItemImageText
oitem.FRectItemId	= itemid
oitem.GetItemImageTextOne

%>
<script language="javascript">

function jsSubmitModi(frm) {
	if (frm.modifiedtext.value == "") {
		alert("������ �Է��ϼ���.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="itemImageText_Process.asp">
	<input type="hidden" name="mode" value="modi">
	<input type="hidden" name="itemid" value="<%= itemid %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="100">��ǰ�ڵ�</td>
		<td bgcolor="#FFFFFF" colspan="3"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>" target="_blank"><%= oitem.FOneItem.Fitemid %></a></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td width="100">��ǰ��</td>
		<td bgcolor="#FFFFFF" colspan="3"><a href="http://www.10x10.co.kr/shopping/category_prd.asp?itemid=<%= oitem.FOneItem.Fitemid %>" target="_blank"><%= oitem.FOneItem.FitemName %></a></td>
    </tr>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td colspan="2">����� �ؽ�Ʈ</td>
		<td colspan="2">���� �ؽ�Ʈ</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td colspan="2">
			<textarea class="textarea_ro" name="imagetext" cols="80" rows="40" readonly><%= oitem.FOneItem.Fimagetext %></textarea>
		</td>
		<td colspan="2">
			<textarea class="textarea" name="modifiedtext" cols="80" rows="40"><%= oitem.FOneItem.Fmodifiedtext %></textarea>
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value="�����ϱ�" onClick="jsSubmitModi(frm)">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
