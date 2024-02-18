<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/yahooitemcls.asp"-->
<%
dim olgeshop
dim page
page = request("page")
if page="" then page=1

dim ix

set olgeshop = new CYahooItemList
olgeshop.FPageSize = 1000
olgeshop.FCurrPage = page
olgeshop.GetLgEshopItem

%>
총 건수 : <%= olgeshop.FtotalCount %> <br>
페이지 : <%= page %>/<%= olgeshop.FtotalPage %><br>
<table width="600" border="1" class="a">
<tr>
	<td>상품ID_Option</td>
	<td>MakerID</td>
	<td>상품명(옵션)</td>
	<td>상품가격</td>
	<td>공급가</td>
</tr>
<% for ix = 0 to olgeshop.FresultCount - 1 %>
<tr>
	<td><%= olgeshop.FItemList(ix).GetItemIdNOption %></td>
	<td><%= olgeshop.FItemList(ix).FMakerID %></td>
	<td><%= olgeshop.FItemList(ix).GetItemIdNOptionName %></td>
	<td><%= olgeshop.FItemList(ix).FSellCash %></td>
	<td><%= CLng(olgeshop.FItemList(ix).FSellCash * 0.81) %></td>
</tr>
<% next %>
</table>
<table width="600" border="0" class="a">
<tr>
	<td colspan="12" align="center">
	<% if olgeshop.HasPreScroll then %>
		<a href="?page=<%= olgeshop.StarScrollPage-1 %>">[pre]</a>
	<% else %>
	<% end if %>

	<% for ix=0 + olgeshop.StarScrollPage to olgeshop.FScrollCount + olgeshop.StarScrollPage - 1 %>
		<% if ix > olgeshop.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(ix) then %>
		<font color="red">[<%= ix %>]</font>
		<% else %>
		<a href="?page=<%= ix %>">[<%= ix %>]</a>
		<% end if %>
	<% next %>

	<% if olgeshop.HasNextScroll then %>
		<a href="?page=<%= ix %>">[next]</a>
	<% else %>
	<% end if %>
	</td>
</tr>
</table>
<%
set olgeshop = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->