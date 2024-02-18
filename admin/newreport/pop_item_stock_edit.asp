<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : �������� ��ǰ ��� ����
' Hieditor : 2011.10.18 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->

<%

dim itemgubun, itemid, itemoption
dim yyyy1, mm1
dim i


itemgubun = requestCheckVar(request("itemgubun"),32)
itemid = requestCheckVar(request("itemid"),32)
itemoption = requestCheckVar(request("itemoption"),32)

yyyy1 = requestCheckVar(request("yyyy1"),32)
mm1 = requestCheckVar(request("mm1"),32)


'// ===========================================================================
dim ostockerrorinfo

set ostockerrorinfo = new CMonthlyStock
    ostockerrorinfo.FRectItemGubun = itemgubun
	ostockerrorinfo.FRectItemId = itemid
	ostockerrorinfo.FRectItemOption = itemoption

	ostockerrorinfo.GetMonthlyErrorInfo
%>

<script language='javascript'>

//����
function UpdateStockMWDiv(frm) {
	if (frm.yyyymm.value == "") {
		alert("������� �����ϼ���");
		return;
	}

	if (frm.lastmwdiv.value == "") {
		alert("���Ա����� �����ϼ���");
		return;
	}

	if (confirm("������Ʈ �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="post" action="pop_item_stock_edit_process.asp">
<input type="hidden" name="mode" value="updatelastmwdiv">
<input type="hidden" name="itemgubun" value="<%= itemgubun %>">
<input type="hidden" name="itemid" value="<%= itemid %>">
<input type="hidden" name="itemoption" value="<%= itemoption %>">

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td width="100" height="25">��ǰ�ڵ�</td>
	<td bgcolor="#FFFFFF">
		<%= itemgubun %>-<%= itemid %>-<%= itemoption %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">���ۿ�</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FMIN_YYYYMM %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">�����</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FMAX_YYYYMM %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">��������</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FMaeipCount %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">��Ź����</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FWitakCount %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">������</td>
	<td bgcolor="#FFFFFF">
		<%= ostockerrorinfo.FOneItem.FErrorCount %>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">���Ա���</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="lastmwdiv">
			<option value="">-����-</option>
			<option value="M">����</option>
			<option value="W">��Ź</option>
		</select>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td height="25">�����</td>
	<td bgcolor="#FFFFFF">
		<select class="select" name="yyyymm">
			<option value="">-����-</option>
			<option value="<%= yyyy1 %>-<%= mm1 %>"><%= yyyy1 %>-<%= mm1 %></option>
			<option value="all">������ ��ü����</option>
		</select>
	</td>
</tr>

<tr bgcolor="<%= adminColor("tabletop") %>">
	<td colspan=2 height="35" align="center">
		<input type="button" class="button" value="������� ������Ʈ" onClick="UpdateStockMWDiv(frm)">
	</td>
</tr>

</form>
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->