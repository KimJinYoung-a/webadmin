<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim makerid, shopid, availstock, research
makerid     = RequestCheckVar(request("makerid"),32)
shopid      = RequestCheckVar(request("shopid"),32)
availstock  = RequestCheckVar(request("availstock"),32)
research    = RequestCheckVar(request("research"),32)

if (research="") and (availstock="") then availstock="on"

dim offstock
set offstock = new COffShopDailyStock
offstock.FRectShopId = shopid
offstock.FRectMakerid = makerid
offstock.FRecAvailStock = availstock

if (makerid<>"") and (shopid<>"") then
	offstock.GetDailyStock
end if

dim i, iptot,retot,selltot,currtot
dim upcheiptot, upcheretot
%>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
			샾 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
			업체:<% drawSelectBoxDesignerwithName "makerid",makerid  %> &nbsp;&nbsp;
			<input type=checkbox name="availstock" <% if availstock="on" then response.write "checked" %> >유효재고만검색
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>
<table width="100%" cellspacing="1" cellpadding="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
    <td width="50">이미지</td>
	<td width="86">바코드</td>
	<td width="100">상품명</td>
	<td width="70">옵션명</td>
	<td width="50">판매가</td>
	<td width="50">온라인가격</td>
	<td width="70">이전<br>실사일</td>
	<td width="40">이전<br>실사</td>
	<td width="40">물류<br>입고</td>
	<td width="40">업체<br>입고</td>
	<td width="40">물류<br>반품</td>
	<td width="40">업체<br>반품</td>
	<td width="40">판매량</td>
	<td width="40">예상재고</td>
</tr>
<% for i=0 to offstock.FresultCount-1 %>
<%
	iptot = iptot + offstock.FItemList(i).Fipno
	upcheiptot = upcheiptot + offstock.FItemList(i).Fupcheipno
	retot = retot + offstock.FItemList(i).Freno
	upcheretot = upcheretot + offstock.FItemList(i).Fupchereno
	selltot = selltot + offstock.FItemList(i).Fsellno
	currtot = currtot + offstock.FItemList(i).Fcurrno
%>
<tr bgcolor="#FFFFFF">
	<td><img src="<%= offstock.FItemList(i).Fimgsmall %>" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
	<td><%= offstock.FItemList(i).GetBarCode %></td>
	<td><%= offstock.FItemList(i).FItemName %></td>
	<td><%= offstock.FItemList(i).FItemOptionName %></td>
	<td align=right>
	<% if offstock.FItemList(i).Fitemgubun="10" and offstock.FItemList(i).Fshopitemprice<>offstock.FItemList(i).Fonlinesellcash then %>
	<font color=red><%= formatNumber(offstock.FItemList(i).Fshopitemprice,0) %></font>
	<% else %>
	<%= formatNumber(offstock.FItemList(i).Fshopitemprice,0) %>
	<% end if %>
	</td>
	<td align=right><%= formatNumber(offstock.FItemList(i).Fonlinesellcash,0) %></td>
	<td align="center"><%= offstock.FItemList(i).Flastrealdate %></td>
	<td align="center"><%= offstock.FItemList(i).Flastrealno %></td>
	<td align="center"><%= offstock.FItemList(i).Fipno %></td>
	<td align="center"><%= offstock.FItemList(i).Fupcheipno %></td>
	<td align="center"><%= offstock.FItemList(i).Freno %></td>
	<td align="center"><%= offstock.FItemList(i).Fupchereno %></td>
	<td align="center"><%= offstock.FItemList(i).Fsellno %></td>
	<% if offstock.FItemList(i).Fcurrno<1 then %>
	<td align="center"><font color="red"><b><%= offstock.FItemList(i).Fcurrno %></font></b></td>
	<% else %>
	<td align="center"><%= offstock.FItemList(i).Fcurrno %></td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="7">total</td>
	<td align="center"></td>
	<td align="center"><%= iptot %></td>
	<td align="center"><%= upcheiptot %></td>
	<td align="center"><%= retot %></td>
	<td align="center"><%= upcheretot %></td>
	<td align="center"><%= selltot %></td>
	<td align="center"><%= currtot %></td>
</tr>
</table>
<%
set offstock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->