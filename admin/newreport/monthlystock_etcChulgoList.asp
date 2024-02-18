<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
dim yyyy1,mm1,sysorreal,research
dim mwgubun,vatyn,minusinc,bPriceGbn,buseo,itemgubun, showsuply, makerid, purchasetype, stplace
dim socid, ipchulcode
dim shopid, chulgogubun
dim grptype, showItemGubun, showIpchulCode

yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)
sysorreal   = requestCheckvar(request("sysorreal"),10)
research    = requestCheckvar(request("research"),10)

mwgubun     = requestCheckvar(request("mwgubun"),10)
vatyn       = requestCheckvar(request("vatyn"),10)
minusinc   = requestCheckvar(request("minusinc"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
buseo       = requestCheckvar(request("buseo"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
showsuply   = requestCheckvar(request("showsuply"),10)
makerid     = requestCheckvar(request("makerid"),32)
purchasetype   = requestCheckvar(request("purchasetype"),10)
stplace     = requestCheckvar(request("stplace"),10)
socid       = requestCheckvar(request("socid"),32)
ipchulcode  = requestCheckvar(request("ipchulcode"),10)
shopid      = requestCheckvar(request("shopid"),32)
chulgogubun = requestCheckvar(request("chulgogubun"),32)

showItemGubun = requestCheckvar(request("showItemGubun"),32)
showIpchulCode = requestCheckvar(request("showIpchulCode"),32)

if (sysorreal="") then sysorreal="sys"  ''real

if (bPriceGbn="") then
    bPriceGbn="P"
end if

if (stplace="") then
    stplace="L"
end if


if (showItemGubun = "Y") and (socid <> "") then
	grptype = "itemgubun"
end if
if (showIpchulCode = "Y") and (socid = "") then
	grptype = "ipchulcode"
end if


''stplace="L"

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


dim ojaego
set ojaego = new CMonthlyStock
ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
ojaego.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
ojaego.FRectTargetGbn = buseo
ojaego.FRectMwDiv    = mwgubun
ojaego.FRectVatYn    = vatyn
ojaego.FRectItemGubun = itemgubun
ojaego.FRectShopSuplyPrice = CHKIIF(showsuply="on",1,0)
ojaego.FRectPlaceGubun = stplace
''ojaego.FRectPurchaseType = purchasetype
ojaego.FRectShopID    = shopid

ojaego.FRectPriceGubun = bPriceGbn
ojaego.FRectChulgoGubun = chulgogubun

ojaego.FRectMakerid = makerid
ojaego.FRectSocID = socid
ojaego.FRectIpChulCode = ipchulcode
ojaego.FRectGrpType = grptype
''if (ojaego.FRectPlaceGubun="L") then '' 오프 일단 제외.. 느림.
    ojaego.GetMonthlyEtcChulgoList
''end if
dim i

dim sumFTTLCNT,sumFTTLSellSum,sumFTTLSuplySum,sumFTTLBuySum,sumFTTLMayStockPrice
dim sumFTTLBuySumMaeipLedger

'' FMayStockPrice =>TTLBuySumAvg??
%>
<script language='javascript'>
function fnResearch(compname,compval,compname2,compval2){
    var frm = document.frm;
    eval('document.frm.'+compname).value=compval;
    eval('document.frm.'+compname2).value=compval2;
    frm.submit();
}

function EdtChulgo(imasterIdx){
    var popURL = '/admin/newstorage/chulgodetail.asp?idx='+imasterIdx+'&menupos=540'
    var popwin=window.open(popURL,'EdtChulgo','width=900,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function EdtIplgo(imasterIdx){
    var popURL = '/common/offshop/shop_ipchuldetail.asp?menupos=196&idx=' + imasterIdx
    var popwin=window.open(popURL,'EdtIplgo','width=900,height=800,scrollbars=yes,resizable=yes');
    popwin.focus();
}

//재고 pop
function TnPopItemStockWithGubun(itemgubun,itemid,itemoption){
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"jsSearchItemStock","width=1200 height=600 scrollbars=yes resizable=yes");

	popwin.focus();
}

//2016/05/09 추가
function PopItemStockShop(shopid, itemgubun, itemid, itemoption) {
	var barcode, formatLength;
	if (itemid*1 >= 1000000) {
		formatLength = 8;
	} else {
		formatLength = 6;
	}

	while (itemid.length < formatLength) {
		itemid = "0" + itemid;
	}

	barcode = itemgubun + itemid + itemoption;

	var popwin = window.open("http://webadmin.10x10.co.kr/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=" + shopid + "&barcode=" + barcode,"PopItemStockShop","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function jsPopApplyAvgPriceToEtcChulgo() {
	var popwin = window.open("monthlystock_etcChulgoList_process.asp?mode=etcavgprc&yyyymm=<%= (yyyy1 + "-" + mm1) %>","jsPopApplyAvgPriceToEtcChulgo","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="" target="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 말일자 재고자산

	        	&nbsp;&nbsp;|&nbsp;&nbsp;

	        	과세구분
	        	<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >전체
	        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >과세
	        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >면세

	        	&nbsp;&nbsp;<input type="checkbox" name="showsuply" value="on" <%= CHKIIF(showsuply="on","checked","") %> >공급가로 표시

                &nbsp;&nbsp;|&nbsp;&nbsp;
                브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
                &nbsp;&nbsp;|&nbsp;&nbsp;
                출고처 : <input type="text" class="text" name="socid" value="<%=socid%>" size="20" >
                &nbsp;&nbsp;|&nbsp;&nbsp;
                출고코드 : <input type="text" class="text" name="ipchulcode" value="<%=ipchulcode%>" size="10" >
		</td>

		<td rowspan="5" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.target='';document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">재고구분:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >시스템재고
        	<!--
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >실사재고
        	-->
        	&nbsp;&nbsp;

        	<font color="#CC3333">매입구분:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >전체
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >매입(+출고매입)
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >위탁
        	<!-- <input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >업체 -->
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >미지정

            <input type="radio" name="mwgubun" value="B012" <% if mwgubun="B012" then response.write "checked" %> >업체위탁
            <input type="radio" name="mwgubun" value="B013" <% if mwgubun="B013" then response.write "checked" %> >출고위탁
            <input type="radio" name="mwgubun" value="B022" <% if mwgubun="B022" then response.write "checked" %> >매장매입
            <input type="radio" name="mwgubun" value="B031" <% if mwgubun="B031" then response.write "checked" %> >출고매입
			<input type="radio" name="mwgubun" value="MWC" <% if mwgubun="MWC" then response.write "checked" %> >매입(+출고매입+위탁)
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		<font color="#CC3333">마이너스구분:</font>
		<input type="radio" name="minusinc" value="" <%= CHKIIF(minusinc="","checked","") %> >마이너스재고 포함(전체)
		<!--
		<input type="radio" name="minusinc" value="P" <%= CHKIIF(minusinc="P","checked","") %> >(+)재고만
	    <input type="radio" name="minusinc" value="M" <%= CHKIIF(minusinc="M","checked","") %> >마이너스재고 만
	    -->
	    &nbsp;&nbsp;
	    <font color="#CC3333">매입가기준:</font>
	    <input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >작성시매입가
		<input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %>  >평균매입가
	    <!--
	    <input type="radio" name="bPriceGbn" value="" <%= CHKIIF(bPriceGbn="","checked","") %>  >현재매입가
	    -->
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <font color="#CC3333">재고위치:</font>
		    <select name="stplace">
        	<option value="L" <%= CHKIIF(stplace="L","selected" ,"") %> >물류
        	<option value="S" <%= CHKIIF(stplace="S","selected" ,"") %> >매장
        	</select>
		    &nbsp;&nbsp;&nbsp;
	    	<font color="#CC3333">부서구분:</font>
	        <select name="buseo">
			<option value="" <%= CHKIIF(buseo="","selected" ,"") %> >전체
			<option value="3X" <%= CHKIIF(buseo="3X","selected" ,"") %> >텐바이텐(3PL제외)
        	<option value="ON" <%= CHKIIF(buseo="ON","selected" ,"") %> >온라인
        	<option value="OF" <%= CHKIIF(buseo="OF","selected" ,"") %> >오프라인
			<option value="IT" <%= CHKIIF(buseo="IT","selected" ,"") %> >아이띵소(구)
			<option value="ET" <%= CHKIIF(buseo="ET","selected" ,"") %> >3PL(아이띵소)
			<option value="EG" <%= CHKIIF(buseo="EG","selected" ,"") %> >3PL(유그레잇)
        	</select>
			&nbsp;&nbsp;&nbsp;
	    	<font color="#CC3333">상품구분:</font>
        	<select name="itemgubun">
        	<option value="" <%= CHKIIF(itemgubun="","selected" ,"") %> >전체
        		<option value="10" <%= CHKIIF(itemgubun="10","selected" ,"") %> >일반(10)</option>
        		<option value="55" <%= CHKIIF(itemgubun="55","selected" ,"") %> >기타(55)</option>
				<option value="70" <%= CHKIIF(itemgubun="70","selected" ,"") %> >소포품(70)</option>
				<option value="75" <%= CHKIIF(itemgubun="75","selected" ,"") %> >저장품(75)</option>
        		<option value="85" <%= CHKIIF(itemgubun="85","selected" ,"") %> >사은품(85)</option>
        		<option value="80" <%= CHKIIF(itemgubun="80","selected" ,"") %> >사은품(80)</option>
        		<option value="90" <%= CHKIIF(itemgubun="90","selected" ,"") %> >오프전용(90)</option>
        	</select>
			&nbsp;&nbsp;&nbsp;
			<input type="checkbox" class="checkbox" name="showItemGubun" value="Y" <%= CHKIIF(showItemGubun = "Y", "checked", "") %> > 상품구분표시
			&nbsp;&nbsp;&nbsp;
			<input type="checkbox" class="checkbox" name="showIpchulCode" value="Y" <%= CHKIIF(showIpchulCode = "Y", "checked", "") %> > 입출코드표시
        	<!--
			&nbsp;&nbsp;&nbsp;
			구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
            -->
			<% if (stplace = "S") then %>
			매장 : <input type="text" class="text" name="shopid" value="<%= shopid %>" size="20" >
			<% end if %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			<font color="#CC3333">출고구분:</font>
			<input type="radio" name="chulgogubun" value="" <% if chulgogubun="" then response.write "checked" %> > 전체
        	<input type="radio" name="chulgogubun" value="etc" <% if chulgogubun="etc" then response.write "checked" %> > 기타출고(etcsales + promotion)
			<input type="radio" name="chulgogubun" value="etc2" <% if chulgogubun="etc2" then response.write "checked" %> > 기타출고(itemgift + itemgift_Biz)
			<input type="radio" name="chulgogubun" value="etc3" <% if chulgogubun="etc3" then response.write "checked" %> > 기타출고(itemgift_all + itemsample + parcelloss + itemdisuse + itemloss + etcout + csservice + shopitemloss + shopitemsample + itemAD) <!-- shopitemloss + shopitemsample 는 매장누적재고 -->
			<!--
			<input type="radio" name="chulgogubun" value="loss" <% if chulgogubun="loss" then response.write "checked" %> > 로스출고(itemgift + itemgift_all + itemsample + parcelloss + itemdisuse + itemloss + etcout + csservice + shopitemloss + shopitemsample)
			-->
			<!--
			<input type="radio" name="chulgogubun" value="cs" <% if chulgogubun="cs" then response.write "checked" %> > CS출고(???)
			-->
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->
<p />

<div>
	<div style="float: left;">*위탁상품+정산안함 제외</div>
	<div style="float: right; margin-bottom: 5px;"><input type="button" class="button" value="평균매입가 적용" onClick="jsPopApplyAvgPriceToEtcChulgo()" /></div>
</div>

<p />

<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td >부서</td>
        <td >출고처</td>
        <td >출고처ID</td>
        <td >구분</td>
        <td >코드구분</td>
        <% if ojaego.FRectListType>="1" then %>
        <td >입출코드</td>
		<td >브랜드</td>
        <% end if %>
        <% if ojaego.FRectListType>"2" then %>
        <td >브랜드</td>
        <td >상품코드</td>
        <% end if %>
    	<td >출고시<br>매입구분</td>
    	<td >재고자산<br>매입구분</td>
    	<td >출고수량</td>
    	<td >출고판매가</td>
    	<td >출고공급가</td>
    	<td >출고매입가</td>
    	<td >차액</td>
		<td >출고매입가(II)</td>
		<!--<td >재고매입가</td>-->

    	<% if ojaego.FRectListType>"1" then %>
			<td >수정</td>
    	<% end if %>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    sumFTTLCNT      = sumFTTLCNT + ojaego.FItemList(i).FTTLCNT
    sumFTTLSellSum  = sumFTTLSellSum + CLNG(ojaego.FItemList(i).FTTLSellSum)
    sumFTTLSuplySum = sumFTTLSuplySum + CLNG(ojaego.FItemList(i).FTTLSuplySum)
    sumFTTLBuySum   = sumFTTLBuySum + CLNG(ojaego.FItemList(i).FTTLBuySum)
    sumFTTLBuySumMaeipLedger = sumFTTLBuySumMaeipLedger + CLNG(ojaego.FItemList(i).FMaeipLedgeravgipgoPrice)
    sumFTTLMayStockPrice   = sumFTTLMayStockPrice + CLNG(ojaego.FItemList(i).FMayStockPrice)

    %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= ojaego.FItemList(i).getBusiName %></td>
        <td><a href="javascript:fnResearch('socid','<%=ojaego.FItemList(i).FSocID%>','buseo','<%=ojaego.FItemList(i).FtargetGbn %>');"><%=ojaego.FItemList(i).FSocName%></a></td>
        <td><%=ojaego.FItemList(i).FSocID%></td>
        <td><%=ojaego.FItemList(i).getITemGubunName%></td>
        <td><%=ojaego.FItemList(i).FItemgubun%></td>
        <% if ojaego.FRectListType>="1" then %>
        <td ><a href="javascript:fnResearch('ipchulcode','<%=ojaego.FItemList(i).FIpChulCode%>','buseo','<%=ojaego.FItemList(i).FtargetGbn %>');"><%=ojaego.FItemList(i).FIpChulCode%></a></td>
		<td ><%=ojaego.FItemList(i).FMakerid%></td>
        <% end if %>
        <% if ojaego.FRectListType>"2" then %>
        <td ><%=ojaego.FItemList(i).FMakerid%></td>
            <% if (shopid<>"") then %>
            	<td>
            		<a href="javascript:PopItemStockShop('<%=shopid%>','<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>')"><%= ojaego.FItemList(i).getLogisticsCode %></a>
            	</td>
            <% else %>
            	<td>
            		<a href="javascript:TnPopItemStockWithGubun('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>')"><%= ojaego.FItemList(i).getLogisticsCode %></a>
            	</td>
            <% end if %>
        <% end if %>
        <td>
        <% if ojaego.FItemList(i).FIpChulMwGubun<>ojaego.FItemList(i).Flastmwdiv then %>
       <font color=red><%=ojaego.FItemList(i).FIpChulMwGubun%></font>
        <% else %>
        <%=ojaego.FItemList(i).FIpChulMwGubun%>
        <% end if %>
        </td>
        <td><%=ojaego.FItemList(i).Flastmwdiv%></td>
        <td><%=ojaego.FItemList(i).FTTLCNT%></td>
        <td align="right"><%=FormatNumber(ojaego.FItemList(i).FTTLSellSum,0)%></td>
        <td align="right"><%=FormatNumber(ojaego.FItemList(i).FTTLSuplySum,0)%></td>
        <td align="right"><%=FormatNumber(ojaego.FItemList(i).FTTLBuySum,0)%></td>
        <td align="right">
	        <% if ojaego.FItemList(i).FMayStockPrice<>ojaego.FItemList(i).FTTLBuySum then %>
	    	    <font color=red><%=FormatNumber(ojaego.FItemList(i).FMayStockPrice-ojaego.FItemList(i).FTTLBuySum,0)%></font>
	        <% end if %>
        </td>
        <td align="right"><%=FormatNumber(ojaego.FItemList(i).FMaeipLedgeravgipgoPrice,0)%></td>
        <!--<td align="right"><%'=FormatNumber(ojaego.FItemList(i).FMayStockPrice,0)%></td>-->

        <% if ojaego.FRectListType>"1" then %>
	    	<td >
				<% if (ojaego.FItemList(i).FSocID = "shopitemloss" or ojaego.FItemList(i).FSocID = "shopitemsample") then %>
				<a href="javascript:EdtIplgo(<%=CLNG(Right(ojaego.FItemList(i).FIpChulCode,6))%>);">수정</a>
				<% else %>
				<a href="javascript:EdtChulgo(<%=CLNG(Right(ojaego.FItemList(i).FIpChulCode,6))%>);">수정</a>
				<% end if %>
			</td>
    	<% end if %>
    </tr>
    <% next %>
    <tr align="center" bgcolor="#FFFFFF">
        <td>합계</td>
        <td></td>
        <td></td>
        <td></td>
        <% if ojaego.FRectListType>="1" then %>
        <td></td>
		<td></td>
        <% end if %>
        <% if ojaego.FRectListType>"2" then %>
        <td></td>
        <td></td>
        <% end if %>
        <td></td>
        <td></td>
        <td></td>

        <td><%=sumFTTLCNT%></td>
        <td align="right"><%=FormatNumber(sumFTTLSellSum,0)%></td>
        <td align="right"><%=FormatNumber(sumFTTLSuplySum,0)%></td>
        <td align="right"><%=FormatNumber(sumFTTLBuySum,0)%></td>
        <td align="right">
	        <% if sumFTTLMayStockPrice<>sumFTTLBuySum then %>
	    	    <font color=red><%=FormatNumber(sumFTTLMayStockPrice-sumFTTLBuySum,0)%></font>
	        <% end if %>
        </td>
        <td align="right"><%=FormatNumber(sumFTTLBuySumMaeipLedger,0)%></td>
        <!--<td align="right"><% '=FormatNumber(sumFTTLMayStockPrice,0)%></td>-->

        <% if ojaego.FRectListType>"1" then %>
    		<td></td>
    	<% end if %>
    </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
