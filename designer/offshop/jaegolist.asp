<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim makerid, shopid, availstock, research, onlyusing
makerid = session("ssBctID")
shopid = request("shopid")
availstock = request("availstock")
onlyusing = request("onlyusing")
research = request("research")

if (research="") and (availstock="") then availstock="on"
if (research="") and (onlyusing="") then onlyusing="on"

dim offstock
set offstock = new COffShopDailyStock
offstock.FRectShopId = shopid
offstock.FRectMakerid = makerid
offstock.FRecAvailStock = availstock
offstock.FRecOnlyusing = onlyusing

if (makerid<>"") and (shopid<>"") then
	offstock.GetDailyStock
end if

dim i, iptot,retot,selltot,currtot

%>
<script language='javascript'>
function searchJ(frm){
	if (frm.shopid.value.length<1){
		alert('샾아이디를 선택하세요.');
		return;
		frm.shopid.focus();
	}
	frm.submit();
}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			SHOP : <% drawSelectBoxOpenOffShop "shopid",shopid %>
			&nbsp;
			<input type=checkbox name="availstock" <% if availstock="on" then response.write "checked" %> >유효재고만검색
			&nbsp;
			<input type=checkbox name="onlyusing" <% if onlyusing="on" then response.write "checked" %> >사용상품만검색
		</td>
		
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:searchJ(frm);">
		</td>
	</tr>
	</form>
</table>

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			<b>예상재고 = 이전실사 + 실사후 입고 - 실사후반품 - 실사후판매</b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="50">이미지</td>
		<td width="86">바코드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="150">이전실사일</td>
		<td width="50">이전<br>실사</td>
		<td width="50">입고</td>
		<td width="50">반품</td>
		<td width="50">판매량</td>
		<td width="50">예상재고</td>
	</tr>
	<% for i=0 to offstock.FresultCount-1 %>
	<%
		iptot = iptot + offstock.FItemList(i).Fipno + offstock.FItemList(i).Fupcheipno
		retot = retot + offstock.FItemList(i).Freno + offstock.FItemList(i).Fupchereno
		selltot = selltot + offstock.FItemList(i).Fsellno
		currtot = currtot + offstock.FItemList(i).Fcurrno
	%>
	<tr bgcolor="#FFFFFF">
		<td><img src="<%= offstock.FItemList(i).Fimgsmall %>" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" width=50 height=50></td>
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
</table>
<%
set offstock = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->