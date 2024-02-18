<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->

<%
dim designer, page, itemid, mode

designer = request("designer")
page = request("page")
itemid = request("itemid")
mode = request("mode")
if page="" then page=1
if mode="" then mode="bybrand"

dim ioffitem
set ioffitem  = new COffShopItem
ioffitem.FPageSize = 50
ioffitem.FCurrPage = page
ioffitem.FRectShopid = "streetshop000"
ioffitem.FRectDesigner = designer

if (itemid<>"") then
	ioffitem.FRectDesigner =  ""
	ioffitem.FRectItemid = itemid
	ioffitem.GetOffLineJumunByItemID
elseif ((mode="bybrand") and (designer<>"")) then
	ioffitem.GetOffLineJumunItem
elseif (mode="byonbest") then
	ioffitem.FRectOrder = "byonbest"
	ioffitem.GetOnlineBestItem
elseif (mode="byonfav") then
	ioffitem.FRectOrder = "byonfav"
	ioffitem.GetOnlineBestItem
elseif (mode="byoffbest") then
	ioffitem.GetOffLineBestItem
elseif (mode="byrecent") then
	ioffitem.FRectOrder = "byrecent"
	ioffitem.GetOffLineJumunItem
elseif (mode="byetc") then
	ioffitem.FRectOrder = "byetc"
	ioffitem.GetOffLineJumunItem
elseif (mode="byevent") then
	ioffitem.FRectOrder = "byevent"
	ioffitem.GetOffLineJumunItem
elseif (mode="onlyoffline") then
	ioffitem.GetOffLineItemList
end if

dim i
%>
<script language='javascript'>
function popOffItemEdit(ibarcode){
	var popwin = window.open('popoffitemedit.asp?barcode=' + ibarcode,'offitemedit','width=500,height=800,resizable=yes,scrollbars=yes');
	popwin.focus();
}

function popOffImageEdit(ibarcode){
	var popwin = window.open('popoffimageedit.asp?barcode=' + ibarcode,'popoffimageedit','width=500,height=600,resizable=yes,scrollbars=yes');
	popwin.focus();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr>
		<td class="a" >
			검색타입 :
			<input type="radio" name="mode" value="bybrand" <% if mode="bybrand" then response.write "checked" %> >브랜드
			<input type="radio" name="mode" value="byonbest" <% if mode="byonbest" then response.write "checked" %> >온라인Best
			<input type="radio" name="mode" value="byonfav" <% if mode="byonfav" then response.write "checked" %> >온라인Favorate
			<input type="radio" name="mode" value="byoffbest" <% if mode="byoffbest" then response.write "checked" %> >오프라인Best
			<input type="radio" name="mode" value="byrecent" <% if mode="byrecent" then response.write "checked" %> >신상품
			<input type="radio" name="mode" value="byevent" <% if mode="byevent" then response.write "checked" %> >행사상품
			<input type="radio" name="mode" value="byetc" <% if mode="byetc" then response.write "checked" %> >기타소모품
			<input type="radio" name="mode" value="onlyoffline" <% if mode="onlyoffline" then response.write "checked" %> >오프라인전용


			<br>
			브랜드 :<% drawSelectBoxDesignerwithName "designer",designer %>
			&nbsp;&nbsp;
			상품코드로검색 : <input type="text" name="itemid" value="<%= itemid %>" size=6 maxlength=7>
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>

<table width="100%" cellspacing="1" class="a" bgcolor=#3d3d3d>
	<% if ioffitem.FresultCount>0 then %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="right">총건수: <%= ioffitem.FTotalCount %> &nbsp; <%= Page %>/<%= ioffitem.FTotalPage %></td>
	</tr>
	<% end if %>
	<tr bgcolor="#DDDDFF">
		<td width="20"></td>
		<td width="50">이미지</td>
		<td width="50">브랜드ID</td>
		<td width="80">BarCode</td>
		<td width="100">상품명</td>
		<td width="80">옵션명</td>
		<td width="60">판매가</td>
		<td width="60">직영공급가</td>
		<td width="48">공급마진</td>
		<td width="70">비고</td>
	</tr>
	<% for i=0 to ioffitem.FResultCount -1 %>
	<tr bgcolor="#FFFFFF">
		<td ><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<% if ioffitem.FItemList(i).FItemGubun<>"10" then %>
		<td ><a href="javascript:popOffImageEdit('<%= ioffitem.FItemList(i).GetBarCode %>')"><img src="<%= ioffitem.FItemList(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" border=0></a></td>
		<% else %>
		<td ><img src="<%= ioffitem.FItemList(i).GetImageSmall %>" width=50 height=50 onError="this.src='http://image.10x10.co.kr/images/no_image.gif'"></td>
		<% end if %>
		<td ><%= ioffitem.FItemList(i).FMakerid %></td>
		<td ><font color="<%= ioffitem.FItemList(i).getSoldOutColor %>"><a href="javascript:popOffItemEdit('<%= ioffitem.FItemList(i).GetBarCode %>');"><%= ioffitem.FItemList(i).GetBarCode %></a></font></td>
		<td ><font color="<%= ioffitem.FItemList(i).getSoldOutColor %>"><%= ioffitem.FItemList(i).FShopItemName %></font></td>
		<td ><font color="<%= ioffitem.FItemList(i).getSoldOutColor %>"><%= ioffitem.FItemList(i).FShopItemOptionName %></font></td>
		<td align=right><%= FormatNumber(ioffitem.FItemList(i).Fshopitemprice,0) %></td>
		<td align=right><%= FormatNumber(ioffitem.FItemList(i).GetOfflineSuplycash,0) %></td>
		<td align=center>
		<% if ioffitem.FItemList(i).Fshopitemprice<>0 then %>
		<%= 100-(CLng(ioffitem.FItemList(i).GetOfflineSuplycash/ioffitem.FItemList(i).Fshopitemprice*10000)/100) %> %
		<% end if %>
		</td>
		<td >
		<% if ioffitem.FItemList(i).Foptusing="N" then %>
		<font color="red">옵션X</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).IsSoldOut then %>
		<font color="red">판매중지</font><br>
		<% end if %>
		<% if ioffitem.FItemList(i).Flimityn="Y" then %>
		<font color="blue">한정(<%= ioffitem.FItemList(i).getLimitNo %>)</font>
		<% end if %>
		</td>
	</tr>
	<% next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="13" align="center">
		<% if ioffitem.HasPreScroll then %>
			<a href="?page=<%= ioffitem.StartScrollPage-1 %>&mode=<%= mode %>&designer=<%= designer %>&research=on">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ioffitem.StartScrollPage to ioffitem.FScrollCount + ioffitem.StartScrollPage - 1 %>
			<% if i>ioffitem.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="?page=<%= i %>&mode=<%= mode %>&designer=<%= designer %>&research=on">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ioffitem.HasNextScroll then %>
			<a href="?page=<%= i %>&mode=<%= mode %>&designer=<%= designer %>&research=on">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<%
set ioffitem  = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->