<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/giftCls.asp"-->

<%
	Dim cGift, vIdx, vGubun, vItemID, vTotSellcash, vSellcash, vDiliItemcost, vUseYN
	vIdx 	= NullFillWith(Request("idx"),"")


	If vIdx <> "" Then
		Set cGift = new ClsGift
		cGift.FIdx = vIdx
		cGift.FGiftCont

		vGubun = cGift.FGubun
		vItemID = cGift.FItemID
		vTotSellcash = cGift.FTot_Sellcash
		vSellcash = cGift.FSellcash
		vDiliItemcost = cGift.FDiliItemcost
		vUseYN = cGift.FUseYN
		set cGift = nothing
	Else
		vUseYN = "Y"
	End If
%>

<script language="javascript">
function jsGofrm()
{
	if(frm1.gubun.value == "")
	{
		alert("������ �����ϼ���.");
		frm1.gubun.focus();
		return false;
	}
	if(frm1.itemid.value == "" || isNaN(frm1.itemid.value))
	{
		alert("��ǰ�ڵ带 ����� �Է��ϼ���.");
		frm1.itemid.focus();
		return false;
	}
	if(frm1.tot_sellcash.value == "" || isNaN(frm1.tot_sellcash.value))
	{
		alert("���ǸŰ��� ����� �Է��ϼ���.");
		frm1.tot_sellcash.focus();
		return false;
	}
	if(frm1.sellcash.value == "" || isNaN(frm1.sellcash.value))
	{
		alert("��ǰ������ ����� �Է��ϼ���.");
		frm1.sellcash.focus();
		return false;
	}
	if(frm1.dili_itemcost.value == "" || isNaN(frm1.dili_itemcost.value))
	{
		alert("��ۺ� ����� �Է��ϼ���.");
		frm1.dili_itemcost.focus();
		return false;
	}
	if(parseInt(frm1.tot_sellcash.value) != parseInt(frm1.sellcash.value) + parseInt(frm1.dili_itemcost.value))
	{
		alert("���ǸŰ��� ��ǰ��+��ۺ� �� Ʋ���ϴ�.");
		return false;
	}
}
function delgift()
{
	if(confirm("�����Ͻ� ��ǰ�� ���� �������� �Ͻðڽ��ϱ�?\n\n���������� �ϸ� �����͸� ���� ����� �˴ϴ�.\n�ݵ�� ������� ���� üũ�ϰ� ���� ���ּ���.") == true) {
		document.frm1.del.value = "o";
		document.frm1.submit();
	}
}
</script>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm1" method="post" action="gift_write_proc.asp" onSubmit="return jsGofrm();">
<input type="hidden" name="idx" value="<%=vIdx%>">
<input type="hidden" name="del" value="x">
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">����</td>
	<td>
		<select name="gubun" <%=CHKIIF(vIdx<>"","disabled","")%>>
			<option value="">-����-</option>
			<option value="giftting" <%=CHKIIF(vGubun="giftting","selected","")%>>������</option>
			<option value="gifticon" <%=CHKIIF(vGubun="gifticon","selected","")%>>����Ƽ��</option>
			<option value="celectory" <%=CHKIIF(vGubun="celectory","selected","")%>>�����丮</option>
			<option value="gsisuper" <%=CHKIIF(vGubun="gsisuper","selected","")%>>GS���̽���</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">��ǰ�ڵ�</td>
	<td>
		<input type="text" name="itemid" value="<%=vItemID%>" maxlength="9" size="10" <%=CHKIIF(vIdx<>"","disabled","")%>>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">���ǸŰ�</td>
	<td>
		<input type="text" name="tot_sellcash" value="<%=vTotSellcash%>" maxlength="7" size="10">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">��ǰ����</td>
	<td>
		<input type="text" name="sellcash" value="<%=vSellcash%>" maxlength="7" size="10">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">��ۺ�</td>
	<td>
		<input type="text" name="dili_itemcost" value="<%=vDiliItemcost%>" maxlength="7" size="10">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="100" bgcolor="<%= adminColor("gray") %>" align="center">��뿩��</td>
	<td>
		<input type="radio" name="useyn" value="Y" <% If vUseYN = "Y" Then %>checked<% End If %>>Y&nbsp;&nbsp;&nbsp;
		<input type="radio" name="useyn" value="N" <% If vUseYN = "N" Then %>checked<% End If %>>N
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td><% If vIdx <> "" Then %><input type="button" class="button" value="��������" onClick="delgift()"><% End If %></td>
	<td align="right">
		<input type="image" src="/images/icon_confirm.gif">
		<a href="javascript:window.close();"><img src="/images/icon_cancel.gif" border="0"></a>
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->