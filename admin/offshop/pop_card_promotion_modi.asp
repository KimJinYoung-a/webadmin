<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ������
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopcardcls.asp"-->
<%

dim idx
idx = requestCheckVar(request("idx"),32)


'// ===========================================================================
dim oOffShopCardPromotion

set oOffShopCardPromotion = new COffShopCardPromotion

oOffShopCardPromotion.FRectIdx = idx

oOffShopCardPromotion.getOneCardPromotion

%>
<script>

function jsCheckValue(frm) {
	if (frm.shopid.value.length < 1) {
		alert('������ �����ϼ���.');
		frm.shopid.focus();
		frm.shopid.select();
		return false;
	}

	if (frm.cardPrice.value.length < 1) {
		alert('����Ʈī��ݾ��� �Է��ϼ���.');
		frm.cardPrice.focus();
		frm.cardPrice.select();
		return false;
	} else if (frm.cardPrice.value*0 != 0) {
		alert('����Ʈī��ݾ��� ���ڸ� �����մϴ�.');
		frm.cardPrice.focus();
		frm.cardPrice.select();
		return false;
	}

	if (frm.startDate.value.length != 10) {
		alert('�������� ��Ȯ�� �Է��ϼ���.');
		frm.startDate.focus();
		frm.startDate.select();
		return false;
	}

	if (frm.endDate.value.length != 10) {
		alert('�������� ��Ȯ�� �Է��ϼ���.');
		frm.endDate.focus();
		frm.endDate.select();
		return false;
	}

	if (frm.rateGubun.value == "") {
		alert('���ݱ����� �����ϼ���.');
		frm.rateGubun.focus();
		return false;
	}

	if (frm.rateAmmount.value.length < 1) {
		alert('���������� �Է��ϼ���.');
		frm.rateAmmount.focus();
		frm.rateAmmount.select();
		return false;
	} else if (frm.rateAmmount.value*0 != 0) {
		alert('���������� ���ڸ� �����մϴ�.');
		frm.rateAmmount.focus();
		frm.rateAmmount.select();
		return false;
	}

	return true;
}

function jsAdd() {
	var frm = document.frm;
	if (jsCheckValue(frm) == true) {
		if (confirm('����Ͻðڽ��ϱ�?')) {
			frm.submit();
		}
	}
}

function jsMody() {
	var frm = document.frm;
	if (jsCheckValue(frm) == true) {
		if (confirm('�����Ͻðڽ��ϱ�?')) {
			frm.submit();
		}
	}
}

</script>
<table width="100%" border="0" cellpadding="2" cellspacing="1" bgcolor="#AAAAAA" class=a>
<form name=frm method=post action="offshop_card_promotion_process.asp">
<input type=hidden name="mode" value="<%= CHKIIF(idx > 0, "modi", "ins") %>">
<input type=hidden name="idx" value="<%= idx %>">
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" width=100>�����̵�</td>
	<td bgcolor="#FFFFFF" >
		<% 'drawSelectBoxOffShop "shopid", oOffShopCardPromotion.FOneItem.Fshopid %>
		<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",oOffShopCardPromotion.FOneItem.Fshopid, "21") %>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" width=100>����Ʈī��ݾ�</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" name="cardPrice" value="<%= oOffShopCardPromotion.FOneItem.FcardPrice %>" size="10" maxlength="10" class="text">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" width=100>������</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" name="startDate" value="<%= oOffShopCardPromotion.FOneItem.FstartDate %>" size="10" maxlength="10" class="text">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" width=100>������</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" name="endDate" value="<%= oOffShopCardPromotion.FOneItem.FendDate %>" size="10" maxlength="10" class="text">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" width=100>���ݱ���</td>
	<td bgcolor="#FFFFFF" >
		<select class="select" name="rateGubun">
			<option></option>
			<option value="1" <%= CHKIIF(oOffShopCardPromotion.FOneItem.FrateGubun=1, "selected", "") %>>����</option>
			<option value="2" <%= CHKIIF(oOffShopCardPromotion.FOneItem.FrateGubun=2, "selected", "") %>>����</option>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" width=100>��������</td>
	<td bgcolor="#FFFFFF" >
		<input type="text" name="rateAmmount" value="<%= oOffShopCardPromotion.FOneItem.FrateAmmount %>" size="10" maxlength="10" class="text">
	</td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>" width=100>��뿩��</td>
	<td bgcolor="#FFFFFF" >
		<select class="select" name="isusing">
			<option value="Y" <%= CHKIIF(oOffShopCardPromotion.FOneItem.Fisusing="Y", "selected", "") %>>Y</option>
			<option value="N" <%= CHKIIF(oOffShopCardPromotion.FOneItem.Fisusing="N", "selected", "") %>>N</option>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan=2 align=center height="35">
		<% if (idx > 0) then %>
		<input type=button class=button value="���θ�� ����" onclick="jsMody()">
		<% else %>
		<input type=button class=button value="���θ�� ���" onclick="jsAdd()">
		<% end if %>

	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
