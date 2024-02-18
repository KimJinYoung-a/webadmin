<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ������ŷ����Ʈ(�ڽ���)
' History : �̻� ����
'			2017.04.11 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/offinvoicecls.asp"-->
<%

menupos = requestCheckVar(request("menupos"),10)

dim page, shopid, statecd
dim research, i

page = requestCheckVar(request("page"),10)
shopid = requestCheckVar(request("shopid"),32)
research = requestCheckVar(request("research"),2)

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
    statecd = "7"
end if

if (page = "") then
	page = 1
end if

dim ocoffinvoice
set ocoffinvoice = new COffInvoice
	ocoffinvoice.FRectShopid = shopid
	ocoffinvoice.FRectStateCD = statecd
	ocoffinvoice.FCurrPage = page
	ocoffinvoice.Fpagesize = 25
	ocoffinvoice.GetMasterList

%>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		ShopID :
		<% if (C_IS_SHOP) then %>
			<%= shopid %>
		<% else %>
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<% end if %>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�˻���� : <b><%= ocoffinvoice.FTotalCount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= ocoffinvoice.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">IDX</td>
	<td>�����̵�</td>
	<td>�κ��̽�<br>NO</td>
	<td>���<br>���</td>
	<td>����<br>�δ�</td>
	<td>����<br>�ñ�</td>
	<td>�ڽ�<br>����</td>
	<td>�ѻ�ǰ�ݾ�<br>(��)</td>
	<td>�ѿ���<br>(��)</td>
	<td>�ۼ�ȭ��</td>
	<td>����ȯ��</td>
	<td>�ѻ�ǰ�ݾ�<br>(��ȯ)</td>
	<td>�ѿ���<br>(��ȯ)</td>
	<td width="80">�����</td>
	<td>���</td>
</tr>
<% if ocoffinvoice.FResultCount >0 then %>
<% for i=0 to ocoffinvoice.FResultcount-1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= ocoffinvoice.FItemList(i).Fidx %></td>
	<td align="center"><a href="shop_offinvoiceview.asp?menupos=<%= menupos %>&idx=<%= ocoffinvoice.FItemList(i).Fidx %>"><%= ocoffinvoice.FItemList(i).Fshopid %><br><%= ocoffinvoice.FItemList(i).Fshopname %></a></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).Finvoiceno %></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).GetDeliverMethodName %></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).GetExportMethodName %></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).GetJungsanTypeName %></td>
	<td align="center"><%= ocoffinvoice.FItemList(i).Ftotalboxno %></td>
	<td align="right">
		<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalgoodsprice, 0) %>&nbsp;
	</td>
	<td align="right">
		<%= FormatNumber(ocoffinvoice.FItemList(i).Ftotalboxprice, 0) %>&nbsp;
	</td>
	<td align="center"><%= ocoffinvoice.FItemList(i).Fpriceunit %></td>
	<td align="center"><%= FormatNumber(ocoffinvoice.FItemList(i).Fexchangerate, 0) %> ��</td>
	<td align="right">
		<% if (ocoffinvoice.FItemList(i).Fexchangerate <> "") and (Not IsNull(ocoffinvoice.FItemList(i).Fexchangerate)) and (ocoffinvoice.FItemList(i).Fexchangerate <> "0") then %>
			<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalgoodsprice/ocoffinvoice.FItemList(i).Fexchangerate), 2), 2) %>&nbsp;
		<% end if %>
	</td>
	<td align="right">
		<% if (ocoffinvoice.FItemList(i).Fexchangerate <> "") and (Not IsNull(ocoffinvoice.FItemList(i).Fexchangerate)) and (ocoffinvoice.FItemList(i).Fexchangerate <> "0") then %>
			<%= FormatNumber(Round((ocoffinvoice.FItemList(i).Ftotalboxprice/ocoffinvoice.FItemList(i).Fexchangerate), 2), 2) %>&nbsp;
		<% end if %>
	</td>
	<td align="center"><%= Left(ocoffinvoice.FItemList(i).Fregdate, 10) %></td>
	<td align="center">
	</td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=15 align=center>[ �˻������ �����ϴ�. ]</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
		<%
		dim strparam
		strparam = "&shopid=" + CStr(shopid)

		strparam = strparam + "&menupos=" + CStr(menupos)

		%>
		<% if ocoffinvoice.HasPreScroll then %>
			<a href="?page=<%= ocoffinvoice.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocoffinvoice.StartScrollPage to ocoffinvoice.FScrollCount + ocoffinvoice.StartScrollPage - 1 %>
			<% if i>ocoffinvoice.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocoffinvoice.HasNextScroll then %>
			<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set ocoffinvoice = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
