<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/shortagestockcls.asp"-->
<%
dim makerid, purchasetype
dim page
page = request("page")
purchasetype = requestCheckVar(request("purchasetype"),32)

if page="" then page=1

dim ostock
set ostock = new CShortageStock
ostock.FCurrPage=page
ostock.FRectPurchaseType = purchasetype
ostock.FPageSize=1000
ostock.GetNoStockList

dim i
%>

<script language='javascript'>
function PopItemSellEdit(iitemid){
	var popwin = window.open('/admin/lib/popitemsellinfo.asp?itemid=' + iitemid,'itemselledit','width=500,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function Research(page){
	document.frm.page.value= page;
	document.frm.submit();
}
</script>


<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			&nbsp;&nbsp;&nbsp;
			구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="20">
			검색결과 : <b><%= ostock.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %> / <%= ostock.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">브랜드</td>
		<td width="50">상품코드</td>
		<td width="50">이미지</td>
		<td>상품명<font color="blue">[옵션명]</font></td>
		<td width="40">마진</td>
		<td width="140">등록일</td>
		<td width="80">비고</td>
	</tr>
<% for i=0 to ostock.FResultCount -1 %>
	<% if ostock.FItemList(i).IsInvalidOption then %>
	<tr align="center" bgcolor="#CCCCCC">
	<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<% end if %>
		<td><a href="/admin/newstorage/orderinput.asp?suplyer=<%= ostock.FItemList(i).FMakerID %>" target="iorderinput"><%= ostock.FItemList(i).FMakerID %></a></td>
		<td><a href="javascript:PopItemSellEdit('<%= ostock.FItemList(i).FItemID %>');"><%= ostock.FItemList(i).FItemID %></a></td>
    	<td width="50" align=center><img src="<%= ostock.FItemList(i).Fimgsmall %>" width=50 height=50></td>
		<td align="left">
			<a href="/admin/stock/itemcurrentstock.asp?itemid=<%= ostock.FItemList(i).FItemID %>&itemoption=<%= ostock.FItemList(i).FItemOption %>" target=_blank ><%= ostock.FItemList(i).FItemName %></a>
			<% if ostock.FItemList(i).FItemOption <> "0000" then %>
				<% if ostock.FItemList(i).Foptionusing="Y" then %>
					<br><font color="blue">[<%= ostock.FItemList(i).FItemOptionName %>]</font>
				<% else %>
					<br><font color="#AAAAAA">[<%= ostock.FItemList(i).FItemOptionName %>]</font>
				<% end if %>
			<% end if %>
		</td>
		<td>
			<font color="<%= ostock.FItemList(i).getMwDivColor %>"><%= ostock.FItemList(i).getMwDivName %></font><br>
			<% if ostock.FItemList(i).Fbuycash<>0 then %>
			<%= 100-(CLng(ostock.FItemList(i).Fbuycash/ostock.FItemList(i).Fsellcash*10000)/100) %> %
			<% end if %>
		</td>
			<td><%= ostock.FItemList(i).Fregdate %></td>
			<td>
			<% if ostock.FItemList(i).Foptionusing="N" then %>
			<font color="red">옵션x</font><br>
			<% end if %>
			<% if ostock.FItemList(i).IsSoldOut then %>
			<font color="red">판매중지</font><br>
			<% end if %>
			<% if ostock.FItemList(i).Flimityn="Y" then %>
			<font color="blue">한정(<%= ostock.FItemList(i).GetLimitStr %>)</font><br>
			<% end if %>
	
			<% if ostock.FItemList(i).Fpreorderno<>0 then %>
			기주문:<%= ostock.FItemList(i).Fpreorderno %>
		<% end if %>
		</td>
	</tr>
<% next %>



<%
set ostock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->