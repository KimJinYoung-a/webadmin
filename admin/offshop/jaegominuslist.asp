<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : 마이너스 재고
' History : 이상구 생성
'			2017.04.13 한용민 수정(보안관련처리)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_dailystock.asp"-->
<%
dim shopid,jaegono, makerid, page
shopid = requestCheckVar(request("shopid"),32)
jaegono = requestCheckVar(request("jaegono"),10)
makerid = requestCheckVar(request("makerid"),32)
page = requestCheckVar(request("page"),10)

if (jaegono="") then jaegono=1
if (page="") then page=1

dim offstock
set offstock = new COffShopDailyStock
offstock.FCurrPage = page
offstock.FPageSize = 100
offstock.FRectMinusNo = jaegono
offstock.FRectMakerid = makerid
offstock.FRectShopId = shopid

if (shopid<>"") then
	offstock.GetCurrentStockMinusList
end if

dim i, iptot,retot,selltot,currtot
%>
<script language='javascript'>
function NextPage(p){
	document.frm.page.value = p;
	document.frm.submit();
}
</script>
<table width="800" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="page" value="">
	<tr>
		<td class="a" >
			샾 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
			업체:<% drawSelectBoxDesignerwithName "makerid",makerid  %>
			<br>
			예상재고
			<input type="text" name="jaegono" value="<%= jaegono %>" size="3" maxlength="4">
			개 미만
		</td>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<table width=800 class=a>
<tr>
	<td align=right>총 <%= offstock.FTotalCount %> 건 <%= page %>/<%= offstock.FtotalPage %> page</td>
</tr>
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
<% if (shopid="") then %>
<tr bgcolor="#FFFFFF">
	<td colspan=10 align=center><font color=red>샾을 선택해 주세요.</font></td>
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
<tr bgcolor="#FFFFFF">
	<td colspan="10" align="center">
	<% if offstock.HasPreScroll then %>
		<a href="javascript:NextPage('<%= offstock.StartScrollPage-1 %>')">[pre]</a>
	<% else %>
		[pre]
	<% end if %>

	<% for i=0 + offstock.StartScrollPage to offstock.FScrollCount + offstock.StartScrollPage - 1 %>
		<% if i>offstock.FTotalpage then Exit for %>
		<% if CStr(page)=CStr(i) then %>
		<font color="red">[<%= i %>]</font>
		<% else %>
		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		<% end if %>
	<% next %>

	<% if offstock.HasNextScroll then %>
		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	<% else %>
		[next]
	<% end if %>
	</td>
</tr>
</table>
<%
set offstock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->