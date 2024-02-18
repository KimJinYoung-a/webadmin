<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/tingcls.asp"-->
<%
dim itemid,itemname
dim page
dim yyyy1,mm1,nowdate
nowdate = Left(CStr(now()),10)

yyyy1 = request("yyyy1")
mm1 = request("mm1")

if (yyyy1="") then
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2) + 1
end if

itemid = request("itemid")
itemname = request("itemname")
page = request("page")

if page="" then page=1

dim iting
set iting = new CTenTenItem
iting.FSearchItemName = itemname
iting.FSearchItemid = itemid
iting.FCurrPage = page
iting.FPageSize = 30
iting.TentenItemList

dim i
%>
<table width="600" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC" align="center">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr>
		<td class="a">
		상품ID:
		<input type="text" name="itemid" value="<%= itemid %>" size="7" maxlength="7" class="input_b">
		상품명:
		<input type="text" name="itemname" value="<%= itemname %>" size="15" maxlength="32" class="input_b">
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="6" align="right" height="30">page: <%= FormatNumber(page,0) %> / <%= FormatNumber(iting.FTotalPage,0) %> count: <%= FormatNumber(iting.FTotalCount,0) %></td>
</tr>
<tr>
	<td align="center">상품ID</td>
	<td align="center">이미지</td>
	<td align="center">상품명</td>
	<td align="center">제안가격</td>
	<td align="center">제안월선택</td>
	<td align="center">제안상품선택</td>
</tr>
<tr>
	<td colspan="6" height="1"><hr noShade color="#DDDDDD" height="1" ></td>
</tr>
<% for i=0 to iting.FresultCount-1 %>
<form name="frmBuyPrc_<%= iting.FItemList(i).FItemID %>" method="post" onSubmit="return false;" action="ting_itemreg_upload.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="itemid" value="<%= iting.FItemList(i).FItemID %>">
<tr height="20">
	<td align="center"><%= iting.FItemList(i).FItemID %></td>
	<td align="center"><img src="<%= iting.FItemList(i).FImageSmall %>" width="50" height="50" border=0 alt=""></td>
	<td align="center"><%= iting.FItemList(i).FItemname %></td>
	<td align="center"><input type="text" name="propcost" size="10" class="input_b"></td>
	<td align="center"><% DrawYMBox yyyy1,mm1 %></td>
	<td align="center"><input type="button" value="선택" onclick="CheckNDoitemviewset(frmBuyPrc_<%= iting.FItemList(i).FItemID %>);"></td>
</tr>
<tr>
	<td colspan="6" height="1"><hr noShade color="#DDDDDD" height="1" ></td>
</tr>
</form>
<% next %>
<tr>
	<td colspan="6" align="center">
	<% if iting.HasPreScroll then %>
		<a href="?page=<%= iting.StarScrollPage-1 %>&itemid=<%= itemid %>&itemname=<%= itemname %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + iting.StarScrollPage to iting.FScrollCount + iting.StarScrollPage - 1 %>
		<% if i>iting.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if iting.HasNextScroll then %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&itemname=<%= itemname %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>

<tr>
	<td colspan="6" height="20">
</tr>
</table>
<%
set iting = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->