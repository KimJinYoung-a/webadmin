<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ۺ���
' Hieditor : 2021.04.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/cscenter/delivery/deliveryTrackCls.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp" -->
<%
dim i, orderserial
    orderserial     = requestCheckVar(request("orderserial"),11)

dim osnap
SET osnap = New COrderMaster
	osnap.FPageSize = 500
	osnap.FCurrPage = 1
    osnap.FRectOrderserial = orderserial
    osnap.getorder_snapshotList()

dim oreward
SET oreward = New COrderMaster
	oreward.FPageSize = 500
	oreward.FCurrPage = 1
    oreward.FRectOrderserial = orderserial
    oreward.getorder_delivery_rewardList()

%>
<script type="text/javascript">

function jsSubmit(frm) {
	frm.submit();
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>" style="margin:0px;">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td  width="50" height="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
        * �ֹ���ȣ : <input type="text" class="text" name="orderserial" value="<%= orderserial %>" size="11" maxlength="11" onKeyPress="if (event.keyCode == 13) jsSubmit(document.frm);">
	</td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:jsSubmit(document.frm);">
	</td>
</tr>
</table>
</form>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
        �� ������� ���� �����
        <br>
        <font color="red">�ֹ� ��ǰ�߿� "�ֹ��ÿ��࿩��" , "�������10���̻󿩺�" �Ѵ� Y �� ��� ����� �Դϴ�.</font>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�ֹ���ȣ</td>
    <td>��ǰ�ڵ�</td>
    <td>�ɼ��ڵ�</td>
    <td>��ǰ��</td>
    <td>�����</td>
    <td>�ֹ��ÿ��࿩��</td>
    <td>�������10���̻󿩺�</td>
</tr>
<% if osnap.FResultCount >0 then %>
<% for i = 0 to osnap.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%= osnap.FItemList(i).forderserial %></td>
    <td><%= osnap.FItemList(i).fitemid %></td>
    <td><%= osnap.FItemList(i).fitemoption %></td>
    <td><%= osnap.FItemList(i).fitemname %></td>
    <td><%= osnap.FItemList(i).fregdt %></td>
    <td><%= osnap.FItemList(i).freserveItemTpyn %></td>
    <td><%= osnap.FItemList(i).fminExpectNoyn %></td>
</tr>
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
        �� ������� : <%= getchulgoscheduledate(orderserial) %>
	</td>
</tr>
</table>
<br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
        �� ������� ���� ���ϸ��� �߱� �α�
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�߱���</td>
    <td>�ֹ���ȣ</td>
    <td>���̵�</td>
    <td>�����</td>
</tr>
<% if oreward.FResultCount >0 then %>
<% for i = 0 to oreward.FResultCount - 1 %>
<tr align="center" bgcolor="#FFFFFF">
    <td><%= oreward.FItemList(i).frewarddate %></td>
    <td><%= oreward.FItemList(i).forderserial %></td>
    <td><%= oreward.FItemList(i).fuserid %></td>
    <td><%= oreward.FItemList(i).fregdt %></td>
</tr>
<% next %>
<% else %>
    <tr bgcolor="#FFFFFF">
        <td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
    </tr>
<% end if %>
</table>
<%
SET osnap = Nothing
SET oreward = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
