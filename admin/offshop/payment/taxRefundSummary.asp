<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� taxRefund ����
' History : 2014.01.17 ������
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/payment/taxRefundMngCls.asp"-->
<%
dim page,shopid,yyyy1,mm1, onlythatdate

shopid = requestCheckvar(request("shopid"),32)
page = requestCheckvar(request("page"),10)
if page="" then page=1
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)
onlythatdate = requestCheckvar(request("onlythatdate"),10)


dim oTaxRefund
set oTaxRefund = new CTaxRefund
%>
<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="A">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">�˻�<br>����</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* ����� :
				&nbsp;&nbsp;
                <input type="checkbox" name="onlythatdate" <%=CHKIIF(onlythatdate="on","checked","") %> >�ش����
                &nbsp;&nbsp;
                &nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if not C_IS_OWN_SHOP and shopid <> "" then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShopAll "shopid",shopid %>
					<% end if %>
				<% else %>
					* ���� : <% drawSelectBoxOffShopAll "shopid",shopid %>
				<% end if %>
			</td>
		</tr>
	    </table>
    </td>
	<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
<tr bgcolor="#FFFFFF" >
    <td>
    * �˻����� :

    </td>
</tr>

</form>
</table>

<!-- ǥ ��ܹ� ��-->
<Br>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= oTaxRefund.FTotalCount %></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>�����</td>
	<td>���ſ�</td>
	<td>����</td>
	<td>ī��</td>
	<td>����</td>
	<td>���ϸ���</td>
	<td>��ǰ��</td>
	<td>����Ʈī��</td>
	<td>�հ�</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td >�հ�</td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
	<td align="right"></td>
</tr>

<tr align="center" bgcolor="#FFFFFF">
	<td colspan="15">�˻� ����� �����ϴ�.</td>
</tr>

</table>
<%
set oTaxRefund=Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->