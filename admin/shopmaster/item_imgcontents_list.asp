<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/item_imgcontentscls.asp" -->
<%

Dim oitemadd,ix,page
dim itemid,makerid

page = request("page")
if (page="") then page=1

itemid = request("itemid")
makerid = request("makerid")

'response.write itemid
'response.write makerid

set oitemadd = new CInfoImage
oitemadd.FPageSize = 12
oitemadd.FCurrPage = page
oitemadd.FItemid = itemid
oitemadd.FMakerid = makerid
oitemadd.getInfoImageList

%>
<table cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black" class="a" bgcolor="#FFFFFF">
<form name="frm" method="get" action="/admin/shopmaster/item_imgcontents_list.asp">
<tr>
	<td>상품코드</td>
	<td><input type="text" name="itemid" value="<%= itemid %>" size="11" maxlength="11" onKeyPress="if (event.keyCode == 13) document.frm.submit();"></td>

</tr>
<tr>
	<td>브랜드</td>
	<td><%	drawSelectBoxDesignerWithName "makerid", makerid %></td>
</tr>
<tr>
	<td colspan="2" align="center"><a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22"  border="0"></a></td>
</tr>
</form>
</table>

<table cellpadding="0" cellspacing="0" border="1" align="center" bordercolordark="White" bordercolorlight="black" bgcolor="#FFFFFF">
<tr class="a">
    <td align="center" width="80">이미지</td>
	<td align="center" width="80">Itemid</td>
	<td align="center" width="100">브랜드</td>
	<td align="center" width="200">상품명</td>
</tr>
<% if oitemadd.FResultCount<1 then %>
<tr class="a">
	<td colspan="2" align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
<% for ix =0 to oitemadd.FResultCount -1 %>
<tr class="a">
    <td align="center"><a href="/admin/shopmaster/item_imgcontents_write.asp?mode=edit&itemid=<% =oitemadd.FItemList(ix).Fitemid %>&menupos=<%= menupos %>"><img src="<% =oitemadd.FItemList(ix).Fsmall_img %>" width="50" height="50" border="0"></a></td>
	<td align="center"><% = oitemadd.FItemList(ix).Fitemid %></td>
	<td align="center"><% = oitemadd.FItemList(ix).FMakerid %></td>
	<td align="center"><% = oitemadd.FItemList(ix).Fitemname %></td>
	</tr>
<% next %>
<% end if %>
<tr class="a">
	<td colspan="5" align="center" height="30">
			<% if oitemadd.HasPreScroll then %>
				<a href="?page=<%= oitemadd.StartScrollPage-1 %>&menupos=<%= menupos %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for ix=0 + oitemadd.StartScrollPage to oitemadd.FScrollCount + oitemadd.StartScrollPage - 1 %>
				<% if ix>oitemadd.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(ix) then %>
				<font color="red">[<%= ix %>]</font>
				<% else %>
				<a href="?page=<%= ix %>&menupos=<%= menupos %>">[<%= ix %>]</a>
				<% end if %>
			<% next %>

			<% if oitemadd.HasNextScroll then %>
				<a href="?page=<%= ix %>&menupos=<%= menupos %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
	</td>
</tr>
<tr class="a">
	<td colspan="5" align="center" height="30">
		<a href="item_imgcontents_write.asp?mode=add&menupos=<%= menupos %>">[올리기]</a>
	</td>
</tr>
</table>
<%
set oitemadd = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->