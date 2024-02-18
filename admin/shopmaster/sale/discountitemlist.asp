<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/discountitemcls.asp"-->

<%
dim designerid,itemid
dim malltype, page

designerid = request("designerid")
malltype = request("malltype")
itemid = request("itemid")
page = request("page")

if page="" then page="1"

dim odiscount
set odiscount = new CDiscount
odiscount.FPageSize=30
odiscount.FCurrPage= page
odiscount.FRectMallType = malltype
odiscount.FRectItemID = itemid
odiscount.FRectDesingerID = designerid

odiscount.GetDiscountItemList

dim i

%>
<script language='javascript'>
function orgprice(iitemid){
	var ret = confirm('원가로 변경하시겠습니까?');

	var frm = document.frmorg;
	if (ret){
		frm.itemid.value = iitemid;
		frm.submit();
	}
}
</script>
<table width="95%" border="0" align="center" cellpadding="5" cellspacing="0"  class="a" >
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr>
		<td >
		몰구분 : <% SelectBoxMallDiv "malltype", malltype %>
		디자이너 선택 :
		<% drawSelectBoxDiscountDesigner "designerid",designerid %>
		상품ID :
		<input type="text" name="itemid" value="<%= itemid %>" size="9" maxlength="9">

		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
	<tr>
		<td colspan="2">Total: <%=formatnumber(odiscount.FTotalCount,0)%>
		</td>
	</tr>
</table>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="#3d3d3d">
<tr bgcolor="#DDDDFF">
<!--
	<td align="center" width="20"><input type="checkbox" name="ckall" value=""></td>
-->
	<td align="center" width="40">상품ID</td>
	<td align="center" width="50" >이미지</td>
	<td align="center">상품명</td>
	<td align="center" width="80">디자이너ID</td>
	<td align="center" width="50">현재<br>판매가</td>
	<td align="center" width="50">현재<br>매입가</td>
	<td align="center" width="50">현재<br>마진율</td>

	<td align="center" width="50">원<br>판매가</td>
	<td align="center" width="50">원<br>매입가</td>
	<td align="center" width="50">원<br>마진율</td>

	<td align="center" width="50">할인<br>판매가</td>
	<td align="center" width="50">할인<br>매입가</td>
	<td align="center" width="50">할인<br>마진율</td>

	<td align="center">저장</td>
</tr>
<% for i=0 to odiscount.FResultCount -1 %>
<tr bgcolor="#FFFFFF">
<!--
	<td><input type="checkbox" name="ck" value=""></td>
-->
	<td align="center"><%= odiscount.FItemList(i).FItemID %></td>
	<td><img src="<%= odiscount.FItemList(i).FImageSmall %>" height="50" width="50"></td>
	<td><%= odiscount.FItemList(i).FItemName %></td>
	<td align="center"><%= odiscount.FItemList(i).FMakerID %></td>
	<td align="right"><%= FormatNumber(odiscount.FItemList(i).FSellcash,0) %></td>
	<td align="right"><%= FormatNumber(odiscount.FItemList(i).FBuycash,0) %></td>
	<td align="center">
	<% if odiscount.FItemList(i).FSellcash<>0 then %>
	<%= 100-fix(odiscount.FItemList(i).FBuycash/odiscount.FItemList(i).FSellcash*10000)/100 %>%
	<% end if %>
	</td>
	<td align="right"><%= FormatNumber(odiscount.FItemList(i).Forgprice,0) %></td>
	<td align="right"><%= FormatNumber(odiscount.FItemList(i).Forgsuplycash,0) %></td>
	<td align="center">
	<% if odiscount.FItemList(i).Forgprice<>0 then %>
	<%= 100-fix(odiscount.FItemList(i).Forgsuplycash/odiscount.FItemList(i).Forgprice*10000)/100 %>%
	<% end if %>
	</td>
	<td align="right"><font color="<%= odiscount.FItemList(i).MatchFont(odiscount.FItemList(i).FSellcash,odiscount.FItemList(i).Fsailprice) %>"><%= FormatNumber(odiscount.FItemList(i).Fsailprice,0) %></font></td>
	<td align="right"><font color="<%= odiscount.FItemList(i).MatchFont(odiscount.FItemList(i).FBuycash,odiscount.FItemList(i).Fsailsuplycash) %>"><%= FormatNumber(odiscount.FItemList(i).Fsailsuplycash,0) %></font></td>
	<td align="center">
	<% if odiscount.FItemList(i).Forgprice<>0 then %>
	<%= 100-fix(odiscount.FItemList(i).Fsailsuplycash/odiscount.FItemList(i).Fsailprice*10000)/100 %>%
	<% end if %>
	</td>
	<td><input type="button" value="원가로" onClick="orgprice('<%= odiscount.FItemList(i).FItemID %>')"></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="15" align="center">
	<% if odiscount.HasPreScroll then %>
		<a href="?page=<%= odiscount.StarScrollPage-1 %>&itemid=<%= itemid %>&malltype=<%= malltype %>&designerid=<%= designerid %>">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + odiscount.StarScrollPage to odiscount.FScrollCount + odiscount.StarScrollPage - 1 %>
		<% if i>odiscount.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&malltype=<%= malltype %>&designerid=<%= designerid %>">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if odiscount.HasNextScroll then %>
		<a href="?page=<%= i %>&itemid=<%= itemid %>&malltype=<%= malltype %>&designerid=<%= designerid %>">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<form name=frmorg method=post action="dodiscountitem.asp">
<input type=hidden name=itemid value="">
</form>
<%
set odiscount = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->