<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jungsan/new_upchejungsancls.asp"-->
<%
dim gubun,idx

gubun   = requestCheckvar(request("gubun"),16)
idx     = requestCheckvar(request("id"),10)


dim ojungsanmaster, IsCommissionTax, itemvatyn
set ojungsanmaster = new CUpcheJungsan
ojungsanmaster.FRectId = idx
ojungsanmaster.JungsanMasterList

if (ojungsanmaster.FResultCount<1) then
    dbget.Close(): response.end
end if

IsCommissionTax = ojungsanmaster.FItemList(0).IsCommissionTax
itemvatyn = ojungsanmaster.FItemList(0).Fitemvatyn
set ojungsanmaster=Nothing

if (IsCommissionTax and (gubun="upche" or gubun="witaksell")) then
    rw "������ �����ΰ�� ��ۺ�, ��Ÿ��� ���길 ��Ÿ�������� ����"
    ''dbget.Close() : response.end
end if
%>

<script language='javascript'>
function adddata(frm){
	if (frm.itemname.value.length<1){
		alert('������ �Է��ϼ���.');
		frm.itemname.focus();
		return;
	}

	if (frm.itemno.value.length<1){
		alert('������ �Է��ϼ���.');
		frm.itemno.focus();
		return;
	}

	if (!IsDigit(frm.itemno.value)){
		alert('������ ���ڸ� �����մϴ�.');
		frm.itemno.focus();
		return;
	}

	if (frm.sellcash.value.length<1){
		alert('�ǸŰ��� �Է��ϼ���.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDigit(frm.sellcash.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.sellcash.focus();
		return;
	}
    <% if (IsCommissionTax) then %>
	if (frm.reducedprice.value.length<1){
		alert('�Ǹ����� �Է��ϼ���.');
		frm.reducedprice.focus();
		return;
	}

	if (!IsInteger(frm.reducedprice.value)){
		alert('�Ǹ����� ���ڸ� �����մϴ�.');
		frm.reducedprice.focus();
		return;
	}

	if (frm.commission.value.length<1){
		alert('�����Ḧ �Է��ϼ���.');
		frm.commission.focus();
		return;
	}

	if (!IsInteger(frm.commission.value)){
		alert('������� ���ڸ� �����մϴ�.');
		frm.commission.focus();
		return;
	}
    <% end if %>

	if (frm.suplycash.value.length<1){
		alert('���԰��� �Է��ϼ���.');
		frm.suplycash.focus();
		return;
	}

	if (!IsInteger(frm.suplycash.value)){
		alert('���԰��� ���ڸ� �����մϴ�.');
		frm.suplycash.focus();
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret){
		frm.submit();
	}
}
</script>


<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>��Ÿ�����߰�</strong></font>
        </td>
        <td align="right">
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frmadd" method="post" action="dodesignerjungsan.asp">
    <input type="hidden" name="mode" value="etcadd">
    <input type="hidden" name="gubun" value="<%= gubun %>">
    <input type="hidden" name="idx" value="<%= idx %>">
    <input type="hidden" name="itemvatyn" value="<%= itemvatyn %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>����</td>
		<td width="40">����</td>
		<td width="80">�ǸŰ�</td>
		<td width="80">�Ǹ���</td>
		<td width="80">������</td>
		<td width="80">���ް�</td>
    </tr>
    <tr bgcolor="#FFFFFF">
		<td><input type="text" name="itemname" value="" size="60"></td>
		<td><input type="text" name="itemno" value="1" size="3" style="text-align:center"></td>
		<td><input type="text" name="sellcash" value="" size="8" style="text-align:right"></td>
		<td><input type="text" name="reducedprice" value="" size="8" style="text-align:right" <%=CHKIIF(NOT IsCommissionTax,"readonly class='text_ro'","")%>></td>
		<td><input type="text" name="commission" value="" size="8" style="text-align:right" <%=CHKIIF(NOT IsCommissionTax,"readonly class='text_ro'","")%>></td>
		<td><input type="text" name="suplycash" value="" size="8" style="text-align:right"></td>
    </tr>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center"><input type="button" value="���� �߰�" onclick="adddata(frmadd)"></td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
    </form>
</table>
<!-- ǥ �ϴܹ� ��-->
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->