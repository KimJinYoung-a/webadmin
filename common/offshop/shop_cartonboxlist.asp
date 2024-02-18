<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  오프라인 샵별패킹리스트(박스별)
' History : 2012.02.02 이상구 생성
'			2012.09.26 한용민 수정
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/common/lib/incMultiLangConst.asp"-->
<!-- #include virtual="/lib/classes/stock/cartoonboxcls.asp"-->
<%
dim page, shopid, showmichulgo, workstate ,research, i
	menupos = requestCheckVar(request("menupos"),10)
	page = requestCheckVar(request("page"),10)
	shopid = requestCheckVar(request("shopid"),32)
	showmichulgo = requestCheckVar(request("showmichulgo"),10)
	research = requestCheckVar(request("research"),2)

if (page = "") then
	page = 1
end if

if (C_IS_SHOP) then
    shopid = C_STREETSHOPID
    workstate = "6,7"
end if

dim occartoonbox
set occartoonbox = new CCartoonBox
	occartoonbox.FRectShopid = shopid
	occartoonbox.FRectShowMichulgo = showmichulgo
	occartoonbox.FRectWorkState = workstate
	occartoonbox.FCurrPage = page
	occartoonbox.Fpagesize = 25
	occartoonbox.GetMasterList

%>

<script type='text/javascript'>

function popSubmaster(iid){
	var popwin = window.open('/offshop/jungsan/franmeaippopsubmaster.asp?idx=' + iid,'popsubmaster','width=800, height=600, scrollbars=yes, resizable=yes');
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>"><%=CTX_SEARCH%><br><%= CTX_conditional %></td>
	<td align="left">
		ShopID :
		<% if (C_IS_SHOP) then %>
			<%= shopid %>
		<% else %>
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
		<% end if %>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="<%=CTX_SEARCH%>" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<br>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9">
		<%= CTX_search_result %> : <b><%= occartoonbox.FTotalCount %></b>
		&nbsp;
		<%= CTX_page %> : <b><%= page %> / <%= occartoonbox.FTotalpage %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">IDX</td>
	<td><%= CTX_title %></td>
	<td><%= CTX_SHOP %></td>
	<td width="80"><%= CTX_Shipment_Date %></td>
	<td width="60"><%= CTX_Invoice_Number %></td>
	<td width="80"><%= CTX_Account %>&nbsp;IDX</td>
	<td width="60"><%= CTX_writer %></td>
	<td width="80"><%= CTX_registration %>&nbsp;<%=CTX_date%></td>
</tr>
<% if occartoonbox.FResultCount >0 then %>
<% for i=0 to occartoonbox.FResultcount-1 %>
<tr bgcolor="#FFFFFF">
	<td align="center"><%= occartoonbox.FItemList(i).Fidx %></td>
	<td align="center">
		<a href="shop_cartonboxview.asp?menupos=<%= menupos %>&idx=<%= occartoonbox.FItemList(i).Fidx %>" onfocus="this.blur()">
		<%= occartoonbox.FItemList(i).Ftitle %></a>
	</td>
	<td align="center">
		<a href="shop_cartonboxview.asp?menupos=<%= menupos %>&idx=<%= occartoonbox.FItemList(i).Fidx %>" onfocus="this.blur()">
		<%= occartoonbox.FItemList(i).Fshopid %><br><%= occartoonbox.FItemList(i).Fshopname %></a>
	</td>
	<td align="center"><%= occartoonbox.FItemList(i).Fdeliverdt %></td>
	<td align="center"><%= occartoonbox.FItemList(i).GetDeliverMethodName %></td>
	<td align="center"><a href="javascript:popSubmaster(<%= occartoonbox.FItemList(i).Fjungsanidx %>)"><%= occartoonbox.FItemList(i).Fjungsanidx %></a></td>
	<td align="center"><%= occartoonbox.FItemList(i).Freguserid %></td>
	<td align="center"><%= Left(occartoonbox.FItemList(i).Fregdate, 10) %></td>
</tr>
<% next %>
<% else %>
<tr bgcolor="#FFFFFF">
	<td colspan=9 align=center>[ 검색결과가 없습니다. ]</td>
</tr>
<% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="9" align="center">
		<%
		dim strparam
		strparam = "&shopid=" + CStr(shopid)

		strparam = strparam + "&menupos=" + CStr(menupos)
		strparam = strparam + "&showmichulgo=" + CStr(showmichulgo)

		%>
		<% if occartoonbox.HasPreScroll then %>
			<a href="?page=<%= occartoonbox.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + occartoonbox.StartScrollPage to occartoonbox.FScrollCount + occartoonbox.StartScrollPage - 1 %>
			<% if i>occartoonbox.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if occartoonbox.HasNextScroll then %>
			<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</table>

<%
set occartoonbox = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
