<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopcostpermeachulcls.asp"-->
<!-- #include virtual="/lib/classes/jungsan/offjungsancls.asp"-->
<%
dim yyyy1,mm1,designer,rectorder, shopid, commcd, minusstockexclude, makerid

yyyy1   			= requestCheckVar(request("yyyy1"),4)
mm1     			= requestCheckVar(request("mm1"),2)
shopid  			= requestCheckVar(request("shopid"),32)
commcd  			= requestCheckVar(request("commcd"),10)
minusstockexclude  	= requestCheckVar(request("minusstockexclude"),10)
makerid  			= requestCheckVar(request("makerid"),32)

if (minusstockexclude = "") then
	minusstockexclude = "N"
end if


'// ===========================================================================
Dim oOffCostDtail
set oOffCostDtail = new COffShopCostPerMeachul
oOffCostDtail.FRectYYYYMM = yyyy1 + "-" + mm1
oOffCostDtail.FRectShopID = shopid
oOffCostDtail.FRectMakerID = makerid
oOffCostDtail.FRectJungsanGubun = commcd
oOffCostDtail.FRectMinusStockExclude = minusstockexclude

oOffCostDtail.GetOffShopCostSumDetailNew


'// ===========================================================================
dim prevmonth

if (yyyy1 <> "") then
	prevmonth = yyyy1 + "-" + mm1 + "-01"
	prevmonth = Left(DateAdd("m", -1, prevmonth), 7)
end if

'// ===========================================================================
Dim i
Dim ttlSell, ttlBuy, ttlChulSum, ttlSuplySum, innerMOrd, innerFOrd, iShopGainSum, iShopCost
Dim StockPricePrevMonth, StockPriceThisMonth
%>
<script language='javascript'>

function popStockShop(yyyy, mm, shopid, makerid, commcd) {
    var popwin=window.open('/admin/newreport/monthlystockShop_detail.asp?menupos=1335&research=on&yyyy1=' + yyyy + '&mm1=' + mm  + '&shopid=' + shopid + '&makerid=' + makerid + '&sysorreal=sys&mwgubun=' + commcd + '&showminus=<% if (minusstockexclude <> "Y") then %>on<% end if %>','popStockShop','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function popJungsanShop(yyyy, mm, shopid, makerid, commcd){
    var popwin=window.open("/admin/offupchejungsan/off_jungsandetailsumBymonth.asp?yyyy1=" + yyyy + "&mm1=" + mm + "&makerid=" + makerid + "&commcd=" + commcd + "&shopid=" + shopid,'popJungsanShop','width=1100,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popChulgoShop(yyyy, mm, shopid, makerid, commcd, iGbn){
    var s    =  new Date();
    s.setFullYear(yyyy,mm*1,0);
    var dd2 = s.getDate();

    var popwin=window.open('/admin/storage/itemipchullist.asp?menupos=168&research=&itemgubun=&itemid=&designer='+makerid+'&gubun=S&yyyy1='+yyyy+'&mm1='+mm+'&dd1=01&yyyy2='+yyyy+'&mm2='+mm+'&dd2='+dd2+'&shopid='+shopid,'popChulgoShop','width=1100,height=600,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popAssignCommCD(imakerid,iyyyymm,ishopid){
    var iURL = "/admin/newreport/popAssignMonthlyCommCd.asp?makerid=" + imakerid+"&yyyymm="+iyyyymm+"&shopid="+ishopid
    var popwin = window.open(iURL,'popAssignMonthlyCommCd','scrollbas=yes,resizable=yes,width=500,height=400');
    popwin.focus();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="rectorder" value="">
	<tr>
		<td class="a" >
			정산대상년월:<% DrawYMBox yyyy1,mm1 %> &nbsp;&nbsp;
			매장 : <% drawSelectBoxOffShop "shopid",shopid %> &nbsp;&nbsp;
			브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %> &nbsp;&nbsp;
		</td>
		<td class="a" align="right" rowspan="2">
			<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	<tr>
		<td class="a" >
			정산구분 : <% drawSelectBoxJungsanCommCombo "commcd",commcd,"Z002" %>
			&nbsp;
			재고구분 :
			<select name="stocktype">
				<option value="SYS">시스템재고</option>
			</select>
			&nbsp;
			마이너스재고 :
			<select name="minusstockexclude">
				<option value="N" <% if (minusstockexclude = "N") then %>selected<% end if %> >포함</option>
				<option value="Y" <% if (minusstockexclude = "Y") then %>selected<% end if %> >제외</option>
			</select>
		</td>
	</tr>
	</form>
</table>
<p>
<table width="100%" border="0" cellpadding="3" cellspacing="1" class="a" bgcolor=#3d3d3d>
<tr bgcolor="#DDDDFF" align="center">
<% if (shopid="") then %>
    <td width="50">매장구분</td>
    <td >매장</td>
<% end if %>
    <td>브랜드ID</td>
    <td width="65">구분(전월)</td>
	<td width="65">정산구분</td>
	<td width="90">총매출<br>(S)</td>
	<td width="10"></td>
	<td width="90">기초재고<br>(B)</td>
	<td width="90">업체정산액<br>(J)</td>
	<td width="90">물류출고액<br>ON매입상품<br>(매입가)<br>(M)</td>
	<td width="90">물류출고액<br>OF매입상품<br>(매입가)<br>(F)</td>
	<td width="90">당월매입합계<br>(A)=J+M+F</td>
	<td width="90">기말재고<br>(E)</td>
	<td width="90">원가<br>(G)=B+A-E</td>
	<td width="80">수익율<br>=(S-G)/S</td>
	<td width="80">재고회전율1<br>=S/(B+E)/2</td>
	<td width="80">재고회전율2<br>=G/(B+E)/2</td>
	<td width="90">물류출고액<br>(매입가)<br>(C)</td>
	<td width="60">기타<br>정산</td>
</tr>
<% for i=0 to oOffCostDtail.FResultCount-1 %>
<%

ttlSell     = ttlSell + oOffCostDtail.FItemList(i).FttlSell
ttlBuy      = ttlBuy + oOffCostDtail.FItemList(i).FttlBuy
ttlChulSum  = ttlChulSum + oOffCostDtail.FItemList(i).FttlChulSum
ttlSuplySum = ttlSuplySum + oOffCostDtail.FItemList(i).FttlSuplySum
innerMOrd   = innerMOrd + oOffCostDtail.FItemList(i).FinnerMOrd
innerFOrd   = innerFOrd + oOffCostDtail.FItemList(i).FinnerFOrd
iShopGainSum = iShopGainSum + oOffCostDtail.FItemList(i).getShopGainSum
iShopCost = iShopCost + oOffCostDtail.FItemList(i).getCostPrice

StockPricePrevMonth     = StockPricePrevMonth + oOffCostDtail.FItemList(i).FstockPricePrevMonth
StockPriceThisMonth     = StockPriceThisMonth + oOffCostDtail.FItemList(i).FstockPriceThisMonth

%>
<tr bgcolor="#FFFFFF" align="right">
<% if (shopid="") then %>
    <td align="center"><%= oOffCostDtail.FItemList(i).FDivName %></td>
    <td align="center"><%= oOffCostDtail.FItemList(i).FShopName %></td>
<% end if %>
    <td align="center"><%= oOffCostDtail.FItemList(i).FMakerid %></td>
    <td align="center">
		<% if (oOffCostDtail.FItemList(i).FComm_cd <> oOffCostDtail.FItemList(i).FPrev_Comm_cd) and (oOffCostDtail.FItemList(i).FstockPricePrevMonth <> 0) then %>
		    <a href="javascript:popAssignCommCD('<%= oOffCostDtail.FItemList(i).FMakerid %>','<%=prevmonth%>','<%= oOffCostDtail.FItemList(i).Fshopid %>')"><font color=red><%= oOffCostDtail.FItemList(i).FPrev_Comm_name %></font></a>
	    <% else %>
	        <a href="javascript:popAssignCommCD('<%= oOffCostDtail.FItemList(i).FMakerid %>','<%=prevmonth%>','<%= oOffCostDtail.FItemList(i).Fshopid %>')"><%= oOffCostDtail.FItemList(i).FPrev_Comm_name %></a>
		<% end if %>
	</td>
    <td align="center">
        <a href="javascript:popAssignCommCD('<%= oOffCostDtail.FItemList(i).FMakerid %>','<%=yyyy1%>-<%=mm1%>','<%= oOffCostDtail.FItemList(i).Fshopid %>')"><%= oOffCostDtail.FItemList(i).FComm_name %></a>
	</td>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCostDtail.FItemList(i).FttlSell,0) %></td>
    <td></td>
	<% if (shopid="") then %>
    	<td bgcolor="#D9D6FF"><a href="javascript:popStockShop('<%= Left(prevmonth, 4) %>', '<%= Right(prevmonth, 2) %>', '<%= oOffCostDtail.FItemList(i).Fshopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '<%= oOffCostDtail.FItemList(i).FPrev_Comm_cd %>')"><%= FormatNumber(oOffCostDtail.FItemList(i).FstockPricePrevMonth,0) %></a></td>
		<td><a href="javascript:popJungsanShop('<%= yyyy1 %>', '<%= mm1 %>', '<%= oOffCostDtail.FItemList(i).Fshopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '')"><%= FormatNumber(oOffCostDtail.FItemList(i).FttlBuy,0) %></a></td>
		<td><a href="javascript:popChulgoShop('<%= yyyy1 %>', '<%= mm1 %>', '<%= oOffCostDtail.FItemList(i).Fshopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '<%= oOffCostDtail.FItemList(i).FComm_cd %>','M')"><%= FormatNumber(oOffCostDtail.FItemList(i).FinnerMOrd,0) %></a></td>
		<td><a href="javascript:popChulgoShop('<%= yyyy1 %>', '<%= mm1 %>', '<%= oOffCostDtail.FItemList(i).Fshopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '<%= oOffCostDtail.FItemList(i).FComm_cd %>','F')"><%= FormatNumber(oOffCostDtail.FItemList(i).FinnerFOrd,0) %></a></td>
		<td bgcolor="#D9D6FF"><%= FormatNumber(oOffCostDtail.FItemList(i).getCostPriceThisMonth,0) %></td>
		<td bgcolor="#D9D6FF"><a href="javascript:popStockShop('<%= yyyy1 %>', '<%= mm1 %>', '<%= oOffCostDtail.FItemList(i).Fshopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '<%= oOffCostDtail.FItemList(i).FComm_cd %>')"><%= FormatNumber(oOffCostDtail.FItemList(i).FstockPriceThisMonth,0) %></a></td>
	<% else %>
    	<td bgcolor="#D9D6FF"><a href="javascript:popStockShop('<%= Left(prevmonth, 4) %>', '<%= Right(prevmonth, 2) %>', '<%= shopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '<%= oOffCostDtail.FItemList(i).FPrev_Comm_cd %>')"><%= FormatNumber(oOffCostDtail.FItemList(i).FstockPricePrevMonth,0) %></a></td>
		<td><a href="javascript:popJungsanShop('<%= yyyy1 %>', '<%= mm1 %>', '<%= shopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '')"><%= FormatNumber(oOffCostDtail.FItemList(i).FttlBuy,0) %></a></td>
		<td><a href="javascript:popChulgoShop('<%= yyyy1 %>', '<%= mm1 %>', '<%= shopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '<%= oOffCostDtail.FItemList(i).FComm_cd %>','M')"><%= FormatNumber(oOffCostDtail.FItemList(i).FinnerMOrd,0) %></a></td>
		<td><a href="javascript:popChulgoShop('<%= yyyy1 %>', '<%= mm1 %>', '<%= shopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '<%= oOffCostDtail.FItemList(i).FComm_cd %>','F')"><%= FormatNumber(oOffCostDtail.FItemList(i).FinnerFOrd,0) %></a></td>
		<td bgcolor="#D9D6FF"><%= FormatNumber(oOffCostDtail.FItemList(i).getCostPriceThisMonth,0) %></td>
		<td bgcolor="#D9D6FF"><a href="javascript:popStockShop('<%= yyyy1 %>', '<%= mm1 %>', '<%= shopid %>', '<%= oOffCostDtail.FItemList(i).FMakerid %>', '<%= oOffCostDtail.FItemList(i).FComm_cd %>')"><%= FormatNumber(oOffCostDtail.FItemList(i).FstockPriceThisMonth,0) %></a></td>
	<% end if %>
    <td bgcolor="#FFD1FF"><%= FormatNumber(oOffCostDtail.FItemList(i).getCostPrice,0) %></td>
    <td><%= oOffCostDtail.FItemList(i).getShopGainPro %></td>
    <td><%= oOffCostDtail.FItemList(i).getTurnoverPro %></td>
    <td><%= oOffCostDtail.FItemList(i).getTurnoverProByCost %></td>
    <td>
        <% if (oOffCostDtail.FItemList(i).isCheckChulSumDiff) then %>
        <%= FormatNumber(oOffCostDtail.FItemList(i).FttlChulSum,0) %>
        <br><font color=red>(<%= FormatNumber(oOffCostDtail.FItemList(i).getChulSumDiffValue,0) %>)</font>
        <% else %>
        <%= FormatNumber(oOffCostDtail.FItemList(i).FttlChulSum,0) %>
        <% end if %>
    </td>
	<td><%= FormatNumber(oOffCostDtail.FItemList(i).FetcBuy,0) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF" align="right">
<% if (shopid="") then %>
    <td align="center" colspan="4">합계</td>
<% else %>
    <td align="center" colspan="2">합계</td>
<% end if %>

    <td></td>
    <td><%= FormatNumber(ttlSell,0) %></td>
    <td></td>
    <td><%= FormatNumber(StockPricePrevMonth,0) %></td>
    <td><%= FormatNumber(ttlBuy,0) %></td>
    <!--
    <td><%= FormatNumber(ttlSuplySum,0) %></td>
    	-->
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
    <td></td>
    <td></td>
    <td><%= FormatNumber(ttlChulSum,0) %></td>
	<td></td>
</tr>
</table>
<%
set oOffCostDtail = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
