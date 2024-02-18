<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/jaego_offline_cls.asp"-->
<%

dim shopid, makerid
shopid = request("shopid")
makerid = request("makerid")


dim ojaegoshop
set ojaegoshop = new CJaegoOffline

ojaegoshop.FRectShopid = shopid
ojaegoshop.FRectMakerid = makerid

ojaegoshop.GetOfflineJeagoSumByShopByMakerByItem


dim i
dim TotalTenMaeipSellPriceSum, TotalTenMaeipBuyPriceSum, TotalTenWitakSellPriceSum, TotalTenWitakBuyPriceSum
dim TotalUpcheWitakSellPriceSum, TotalUpcheWitakBuyPriceSum, TotalUpcheMaeipSellPriceSum, TotalUpcheMaeipBuyPriceSum

TotalTenMaeipSellPriceSum = 0
TotalTenMaeipBuyPriceSum = 0
TotalTenWitakSellPriceSum = 0
TotalTenWitakBuyPriceSum = 0
TotalUpcheWitakSellPriceSum = 0
TotalUpcheWitakBuyPriceSum = 0
TotalUpcheMaeipSellPriceSum = 0
TotalUpcheMaeipBuyPriceSum = 0

dim divname

if (ojaegoshop.FResultCount > 0) then
        divname = ojaegoshop.FItemList(0).getChargeDivName
end if

%>
<script language='javascript'>
function PopShopItemDetail(shopid, itemgubun, itemid, itemoption){
	var popwin = window.open('/admin/stock/itemcurrentstock_shop.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'shopitemdetail','width=1000,height=600,scrollbars=yes,resizable=yes')
}
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#F3F3FF">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif"><img src="/images/icon_star.gif" align="absbottom">
		<font color="red"><strong>SHOP 브랜드별 재고자산</strong></font></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td>
			<br>OFFLINE 에 대한 샆별 브랜드별 현재고 기준 재고자산 정보입니다.(전일 새벽 1시 기준)
		</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10"valign="top">
		<td><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_08.gif"></td>
		<td><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
	</tr>
</table>

<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	샾 &nbsp;: <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
	        	업체: <% drawSelectBoxDesignerwithName "makerid",makerid  %> &nbsp;&nbsp;<br>
	        </td>
	        <td valign="top" align="right">
	        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
    <tr align="center" bgcolor="#DDDDFF">
    	<td width="30">구분</td>
    	<td width="50">상품코드</td>
    	<td>상품명</td>
    	<td width="100">옵션</td>
    	<td width="30">계약<br>조건</td>
    	<td width="50">시스템<br>재고</td>
    	<td width="60">소비자가</td>
    	<td width="60">샆공급가</td>
    	<td width="60">소비자가<br>합계</td>
    	<td width="60">샆공급가<br>합계</td>
    	<td width="100">비고</td>
    </tr>
<%

TotalTenMaeipSellPriceSum = 0
TotalTenMaeipBuyPriceSum = 0
TotalTenWitakSellPriceSum = 0
TotalTenWitakBuyPriceSum = 0
TotalUpcheWitakSellPriceSum = 0
TotalUpcheWitakBuyPriceSum = 0
TotalUpcheMaeipSellPriceSum = 0
TotalUpcheMaeipBuyPriceSum = 0

%>
<% for i=0 to ojaegoshop.FResultCount-1 %>
        <%
        TotalTenMaeipSellPriceSum = TotalTenMaeipSellPriceSum + ojaegoshop.FItemList(i).FTenMaeipSellPriceSum
        TotalTenMaeipBuyPriceSum = TotalTenMaeipBuyPriceSum + ojaegoshop.FItemList(i).FTenMaeipBuyPriceSum
        TotalTenWitakSellPriceSum = TotalTenWitakSellPriceSum + ojaegoshop.FItemList(i).FTenWitakSellPriceSum
        TotalTenWitakBuyPriceSum = TotalTenWitakBuyPriceSum + ojaegoshop.FItemList(i).FTenWitakBuyPriceSum
        TotalUpcheWitakSellPriceSum = TotalUpcheWitakSellPriceSum + ojaegoshop.FItemList(i).FUpcheWitakSellPriceSum
        TotalUpcheWitakBuyPriceSum = TotalUpcheWitakBuyPriceSum + ojaegoshop.FItemList(i).FUpcheWitakBuyPriceSum
        TotalUpcheMaeipSellPriceSum = TotalUpcheMaeipSellPriceSum + ojaegoshop.FItemList(i).FUpcheMaeipSellPriceSum
        TotalUpcheMaeipBuyPriceSum = TotalUpcheMaeipBuyPriceSum + ojaegoshop.FItemList(i).FUpcheMaeipBuyPriceSum
        %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td align="center"><%= ojaegoshop.FItemList(i).Fitemgubun %></td>
    	<td align="center"><%= ojaegoshop.FItemList(i).Fitemid %></td>
    	<td align="left"><a href="javascript:PopShopItemDetail('<%= shopid %>','<%= ojaegoshop.FItemList(i).Fitemgubun %>','<%= ojaegoshop.FItemList(i).Fitemid %>','<%= ojaegoshop.FItemList(i).Fitemoption %>')"><%= ojaegoshop.FItemList(i).Fitemname %></a></td>
    	<td align="left"><%= ojaegoshop.FItemList(i).Fitemoptionname %></td>
    	<td align="center"><font color="<%= ojaegoshop.FItemList(i).getChargeDivColor %>"><%= ojaegoshop.FItemList(i).getChargeDivName %></font></td>
    	<td align="right"><%= ojaegoshop.FItemList(i).Fsysstockno %></td>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).Fshopitemprice,0) %></td>
    	<td align="right"><%= FormatNumber((ojaegoshop.FItemList(i).Fshopitemprice * (1 - (ojaegoshop.FItemList(i).Fdefaultsuplymargin/100))),0) %></td>
<% if (divname = "텐매") then %>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).FTenMaeipSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).FTenMaeipBuyPriceSum,0) %></td>
<% end if %>
<% if (divname = "텐위") then %>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).FTenWitakSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).FTenWitakBuyPriceSum,0) %></td>
<% end if %>
<% if (divname = "업위") then %>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).FUpcheWitakSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).FUpcheWitakBuyPriceSum,0) %></td>
<% end if %>
<% if (divname = "업매") then %>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).FUpcheMaeipSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaegoshop.FItemList(i).FUpcheMaeipBuyPriceSum,0) %></td>
<% end if %>
<% if (divname = "") then %>
    	<td align="right">0</td>
    	<td align="right">0</td>
<% end if %>
    	<td></td>
    </tr>
<% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td colspan="8" align="left">총계</td>
<% if (divname = "텐매") then %>
    	<td align="right"><%= FormatNumber(TotalTenMaeipSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalTenMaeipBuyPriceSum,0) %></td>
<% end if %>
<% if (divname = "텐위") then %>
    	<td align="right"><%= FormatNumber(TotalTenWitakSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalTenWitakBuyPriceSum,0) %></td>
<% end if %>
<% if (divname = "업위") then %>
    	<td align="right"><%= FormatNumber(TotalUpcheWitakSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalUpcheWitakBuyPriceSum,0) %></td>
<% end if %>
<% if (divname = "업매") then %>
    	<td align="right"><%= FormatNumber(TotalUpcheMaeipSellPriceSum,0) %></td>
    	<td align="right"><%= FormatNumber(TotalUpcheMaeipBuyPriceSum,0) %></td>
<% end if %>
<% if (divname = "") then %>
    	<td align="right">0</td>
    	<td align="right">0</td>
<% end if %>
    	<td></td>
    </tr>
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->
<%

set ojaegoshop = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->