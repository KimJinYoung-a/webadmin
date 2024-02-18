<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/offshop/incSessionoffshop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/offshop/lib/offshopbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim makerid, shopid, availstock, research
makerid = request("makerid")
shopid = session("ssBctID")
availstock = request("availstock")
research = request("research")

if (research="") and (availstock="") then availstock="on"

dim offstock
set offstock = new COffShopDailyStock
offstock.FRectShopId = shopid
offstock.FRectMakerid = makerid
offstock.FRecAvailStock = availstock

if (makerid<>"") then
	offstock.GetDailyStock
end if

dim i, iptot,retot,selltot,currtot
%>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr>
		<td class="a" >
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
<table width="800" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align=center>
    <td width="50">이미지</td>
	<td width="86">바코드</td>
	<td width="100">상품명</td>
	<td width="80">옵션명</td>
	<td width="80">이전<br>실사일</td>
	<td width="50">이전<br>실사</td>
	<td width="50">입고</td>
	<td width="50">반품</td>
	<td width="50">판매량</td>
	<td width="50">예상재고</td>
</tr>
<% if (makerid="") then %>
<tr bgcolor="#FFFFFF">
	<td colspan="10" align=center><font color=red>브랜드를 선택해주세요</font></td>
</tr>
<% else %>
<% for i=0 to offstock.FresultCount-1 %>
<%
	iptot = iptot + offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno
	retot = retot + offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno
	selltot = selltot + offstock.FItemList(i).Fsellno
	currtot = currtot + offstock.FItemList(i).Fcurrno
%>
<tr bgcolor="#FFFFFF">
	<td><img src="<%= offstock.FItemList(i).Fimgsmall %>" onError="this.src='http://image.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
	<td><%= offstock.FItemList(i).GetBarCode %></td>
	<td><%= offstock.FItemList(i).FItemName %></td>
	<td><%= offstock.FItemList(i).FItemOptionName %></td>
	<td align="center"><%= offstock.FItemList(i).Flastrealdate %></td>
	<td align="center"><%= offstock.FItemList(i).Flastrealno %></td>
	<td align="center"><%= offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno %></td>
	<td align="center"><%= offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno %></td>
	<td align="center"><%= offstock.FItemList(i).Fsellno %></td>
	<% if offstock.FItemList(i).Fcurrno<1 then %>
	<td align="center"><font color="red"><b><%= offstock.FItemList(i).Fcurrno %></font></b></td>
	<% else %>
	<td align="center"><%= offstock.FItemList(i).Fcurrno %></td>
	<% end if %>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
	<td colspan="5">total</td>
	<td align="center"></td>
	<td align="center"><%= iptot %></td>
	<td align="center"><%= retot %></td>
	<td align="center"><%= selltot %></td>
	<td align="center"><%= currtot %></td>
</tr>
<% end if %>
</table>
<%
set offstock = Nothing
%>
<!-- #include virtual="/offshop/lib/offshopbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->