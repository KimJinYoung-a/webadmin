<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/analysiscls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopcostpermeachulcls.asp"-->

<%
dim yyyy1,mm1,designer,rectorder, shopid, shopdiv, minusstockexclude, makerid
yyyy1   = requestCheckVar(request("yyyy1"),4)
mm1     = requestCheckVar(request("mm1"),2)
shopid  = requestCheckVar(request("shopid"),32)
shopdiv = requestCheckVar(request("shopdiv"),32)
makerid = requestCheckVar(request("makerid"),32)
minusstockexclude = requestCheckVar(request("minusstockexclude"),32)

if (minusstockexclude = "") then
	minusstockexclude = "N"
end if

dim dt
if yyyy1="" then
	dt = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(dt),4)
	mm1 = Mid(CStr(dt),6,2)
end if

'' response.write "TEST"
'' response.end

Dim oOffCost
set oOffCost = new COffShopCostPerMeachul
oOffCost.FRectYYYYMM = yyyy1 + "-" + mm1
oOffCost.FRectShopID=shopid
oOffCost.FRectShopDiv=shopdiv
oOffCost.FRectMakerID=makerid
oOffCost.FRectMinusStockExclude = minusstockexclude
oOffCost.GetOffMainCostSumByShop

Dim oOffCostJs
set oOffCostJs = new COffShopCostPerMeachul
oOffCostJs.FRectYYYYMM = yyyy1 + "-" + mm1
oOffCostJs.FRectShopID=shopid
oOffCostJs.FRectShopDiv=shopdiv
oOffCostJs.FRectMinusStockExclude = minusstockexclude
'oOffCostJs.GetOffMainCostSumByJungsan

dim i
Dim ttlSell, ttlBuy, ttlChulSum, ttlSuplySum, innerMOrd, innerFOrd, iShopGainSum, iShopCost
Dim StockPricePrevMonth, StockPriceThisMonth
%>
<script language='javascript'>
function MakeMonthlyBrandSellSum(yyyymm){
	var popwin=window.open('dooffgainsum.asp?menupos=<%= menupos %>&mode=MakeMonthlyBrandSellSum&yyyymm='+yyyymm,'dooffgainsum','width=100,height=100');
	popwin.focus();
}

function MakeMonthlyBrandStockSum(yyyymm){
	var popwin=window.open('dooffgainsum.asp?menupos=<%= menupos %>&mode=MakeMonthlyBrandStockSum&yyyymm='+yyyymm,'MakeMonthlyBrandStockSum','width=100,height=100');
	popwin.focus();
}

function popOffGainDetail(yyyy,mm,shopid,commcd){
    var popwin=window.open('offgainsumDetail.asp?menupos=<%= menupos %>&yyyy1='+yyyy+'&mm1='+mm+'&shopid='+shopid+'&commcd='+commcd,'offgainsumDetail','width=900,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popOffMainGainDetail(yyyy,mm,shopid){
    var popwin=window.open('offmaingainsumDetail.asp?menupos=<%= menupos %>&yyyy1='+yyyy+'&mm1='+mm+'&shopid='+shopid,'popOffMainGainDetail','width=900,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
		정산대상년월:<% DrawYMBox yyyy1,mm1 %> &nbsp;&nbsp;
		매장 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;

		브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %> &nbsp;&nbsp;

		매장구분 : <% drawSelectBoxShopDiv "shopdiv",shopdiv,true %>
		&nbsp;
		마이너스재고 :
		<select name="minusstockexclude">
			<option value="N" <% if (minusstockexclude = "N") then %>selected<% end if %> >포함</option>
			<option value="Y" <% if (minusstockexclude = "Y") then %>selected<% end if %> >제외</option>
		</select>
		<td class="a" align="right">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table>
<br>

<p>

<font size="5">수정중입니다.</font>

<p>

<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align="center">
    <td width="50">매장구분</td>
    <td >매장</td>
	<td width="90">총매출<br>(공급가)<br>(S)</td>
	<td width="10"></td>
	<td width="90">기초재고<br>(B)</td>
	<td width="90">업체정산액<br>(J)</td>
	<td width="90">물류출고액<br>ON매입상품<br>(매입가)<br>(M)</td>
	<td width="90">물류출고액<br>OF매입상품<br>(매입가)<br>(F)</td>
	<td width="90">당월매입합계<br>(A)=J+M+F</td>
	<td width="90">기말재고<br>(E)</td>
	<td width="90">원가<br>(G)=B+A-E</td>
	<td width="90">수익율<br>=(S-A)/S</td>
	<td width="90">물류출고액<br>(공급가)<br>(C)</td>
	<td width="60">기타보정</td>
</tr>
<% for i=0 to oOffCost.FResultCount-1 %>
<%
ttlSell     = ttlSell + oOffCost.FItemList(i).FttlSell
ttlBuy      = ttlBuy + oOffCost.FItemList(i).FttlBuy
ttlChulSum  = ttlChulSum + oOffCost.FItemList(i).FttlChulSum
ttlSuplySum = ttlSuplySum + oOffCost.FItemList(i).FttlSuplySum
innerMOrd   = innerMOrd + oOffCost.FItemList(i).FinnerMOrd
innerFOrd   = innerFOrd + oOffCost.FItemList(i).FinnerFOrd
iShopGainSum = iShopGainSum + oOffCost.FItemList(i).getShopGainSum
iShopCost = iShopCost + oOffCost.FItemList(i).getCostPrice
StockPricePrevMonth     = StockPricePrevMonth + oOffCost.FItemList(i).FstockPricePrevMonth
StockPriceThisMonth     = StockPriceThisMonth + oOffCost.FItemList(i).FstockPriceThisMonth
%>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center"><%= oOffCost.FItemList(i).FDivName %></td>
    <td align="center"><a href="javascript:popOffMainGainDetail('<%=yyyy1%>','<%=mm1%>','<%= oOffCost.FItemList(i).Fshopid %>');"><%= oOffCost.FItemList(i).FShopName %></a></td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCost.FItemList(i).FttlSell,0) %></td>
    <td></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber(oOffCost.FItemList(i).FstockPricePrevMonth,0) %></td>
    <td>
        <% if (oOffCost.FItemList(i).FShopid="") then %>
        <a href="javascript:popOffGainDetail('<%=yyyy1%>','<%=mm1%>','','B021');"><%= FormatNumber(oOffCost.FItemList(i).FttlBuy,0) %></a>
        <% else %>
        <%= FormatNumber(oOffCost.FItemList(i).FttlBuy,0) %>
        <% end if %>
    </td>
    <td><%= FormatNumber(oOffCost.FItemList(i).FinnerMOrd,0) %></td>
    <td><%= FormatNumber(oOffCost.FItemList(i).FinnerFOrd,0) %></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber((oOffCost.FItemList(i).FttlBuy + oOffCost.FItemList(i).FinnerMOrd + oOffCost.FItemList(i).FinnerFOrd),0) %></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber(oOffCost.FItemList(i).FstockPriceThisMonth,0) %></td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCost.FItemList(i).getCostPrice,0) %></td>
    <td><%= oOffCost.FItemList(i).getShopGainPro %></td>
    <td><%= FormatNumber(oOffCost.FItemList(i).FttlSuplySum,0) %></td>
	<td><%= FormatNumber(oOffCost.FItemList(i).FetcBuy,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center">합계</td>
    <td></td>
    <td><%= FormatNumber(ttlSell,0) %></td>
    <td></td>
    <td><%= FormatNumber(StockPricePrevMonth,0) %></td>
    <td><%= FormatNumber(ttlBuy,0) %></td>
    <td><%= FormatNumber(innerMOrd,0) %></td>
    <td><%= FormatNumber(innerFOrd,0) %></td>
    <td><%= FormatNumber((ttlBuy + innerMOrd + innerFOrd),0) %></td>
    <td><%= FormatNumber(StockPriceThisMonth,0) %></td>
    <td><%= FormatNumber(iShopCost,0) %></td>
    <td>
    <% if ttlSell<>0 then %>
    <%= CLNG(iShopGainSum/ttlSell*100*100)/100 %>
    <% end if %>
    </td>
    <td><%= FormatNumber(ttlSuplySum,0) %></td>
	<td></td>
</tr>
</table>
<p>

<% if False then %>

<%
ttlSell     = 0
ttlBuy      = 0
ttlChulSum  = 0
ttlSuplySum = 0
innerMOrd   = 0
innerFOrd   = 0
iShopGainSum =0
iShopCost	= 0
StockPricePrevMonth = 0
StockPriceThisMonth = 0
%>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align="center">
    <td >구분</td>
	<td width="90">총매출<br>(S)</td>
	<td width="10"></td>
	<td width="90">기초재고<br>(B)</td>
	<td width="90">업체정산액<br>(J)</td>
	<td width="90">물류출고액<br>ON매입상품<br>(매입가)<br>(M)</td>
	<td width="90">물류출고액<br>OF매입상품<br>(매입가)<br>(F)</td>
	<td width="90">당월매입합계<br>(A)=J+M+F</td>
	<td width="90">기말재고<br>(E)</td>
	<td width="90">원가<br>(G)=B+A-E</td>
	<td width="90">수익율<br>=(S-A)/S</td>
	<td width="90">물류출고액<br>(매입가)<br>(C)</td>
	<td width="60">기타보정</td>
</tr>
<% for i=0 to oOffCostJs.FResultCount-1 %>
<%
ttlSell     = ttlSell + oOffCostJs.FItemList(i).FttlSell
ttlBuy      = ttlBuy + oOffCostJs.FItemList(i).FttlBuy
ttlChulSum  = ttlChulSum + oOffCostJs.FItemList(i).FttlChulSum
ttlSuplySum = ttlSuplySum + oOffCostJs.FItemList(i).FttlSuplySum
innerMOrd   = innerMOrd + oOffCostJs.FItemList(i).FinnerMOrd
innerFOrd   = innerFOrd + oOffCostJs.FItemList(i).FinnerFOrd
iShopGainSum = iShopGainSum + oOffCostJs.FItemList(i).getShopGainSum
iShopCost = iShopCost + oOffCostJs.FItemList(i).getCostPrice
StockPricePrevMonth     = StockPricePrevMonth + oOffCostJs.FItemList(i).FstockPricePrevMonth
StockPriceThisMonth     = StockPriceThisMonth + oOffCostJs.FItemList(i).FstockPriceThisMonth
%>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center"><a href="javascript:popOffGainDetail('<%=yyyy1%>','<%=mm1%>','<%=shopid%>','<%= oOffCostJs.FItemList(i).FComm_cd%>');"><%= oOffCostJs.FItemList(i).FComm_name %></a></td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCostJs.FItemList(i).FttlSell,0) %></td>
    <td></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber(oOffCostJs.FItemList(i).FstockPricePrevMonth,0) %></td>
    <td><%= FormatNumber(oOffCostJs.FItemList(i).FttlBuy,0) %></td>
    <td><%= FormatNumber(oOffCostJs.FItemList(i).FinnerMOrd,0) %></td>
    <td><%= FormatNumber(oOffCostJs.FItemList(i).FinnerFOrd,0) %></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber((oOffCostJs.FItemList(i).FttlBuy + oOffCostJs.FItemList(i).FinnerMOrd + oOffCostJs.FItemList(i).FinnerFOrd),0) %></td>
    <td bgcolor="#D9D6FF"><%= FormatNumber(oOffCostJs.FItemList(i).FstockPriceThisMonth,0) %></td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCostJs.FItemList(i).getCostPrice,0) %></td>
    <td><%= oOffCostJs.FItemList(i).getShopGainPro %></td>
    <td>
        <% if (oOffCostJs.FItemList(i).isCheckChulSumDiff) then %>
        <%= FormatNumber(oOffCostJs.FItemList(i).FttlChulSum,0) %>
        <br><font color=red>(<%= FormatNumber(oOffCostJs.FItemList(i).getChulSumDiffValue,0) %>)</font>
        <% else %>
        <%= FormatNumber(oOffCostJs.FItemList(i).FttlChulSum,0) %>
        <% end if %>
    </td>
	<td><%= FormatNumber(oOffCostJs.FItemList(i).FetcBuy,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="right">
    <td align="center">합계</td>
    <td><%= FormatNumber(ttlSell,0) %></td>
    <td></td>
    <td><%= FormatNumber(StockPricePrevMonth,0) %></td>
    <td><%= FormatNumber(ttlBuy,0) %></td>
    <td><%= FormatNumber(innerMOrd,0) %></td>
    <td><%= FormatNumber(innerFOrd,0) %></td>
    <td><%= FormatNumber((ttlBuy + innerMOrd + innerFOrd),0) %></td>
    <td><%= FormatNumber(StockPriceThisMonth,0) %></td>
    <td><%= FormatNumber(iShopCost,0) %></td>
    <td>
    <% if ttlSell<>0 then %>
    <%= CLNG(iShopGainSum/ttlSell*100*100)/100%>
    <% end if %>
    </td>
    <td><%= FormatNumber(ttlChulSum,0) %></td>
	<td></td>
</tr>
</table>
<br>* 계약구분이 변경된 브랜드가 존재할 경우 <font color="red">기초재고금액</font>이 달라집니다.

<% end if %>
<%
set oOffCost = Nothing
set oOffCostJs = Nothing
%>



<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
