<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 재고자산
' History : 이상구 생성
'			2023.05.03 한용민 수정(검색조건추가)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
Const isOnlySys = FALSE
Const isViewWonga =FALSE

dim yyyy1,mm1,isusing,sysorreal, research, newitem, vatyn, minusinc, bPriceGbn, i, brandUseYN
dim mwgubun, buseo, itemgubun, stplace, purchasetype, showsuply, dtype, makerid, shopid, etcjungsantype, showDiff
	yyyy1       = requestCheckvar(request("yyyy1"),10)
	mm1         = requestCheckvar(request("mm1"),10)
	isusing     = requestCheckvar(request("isusing"),10)
	sysorreal   = requestCheckvar(request("sysorreal"),10)
	research    = requestCheckvar(request("research"),10)
	newitem     = requestCheckvar(request("newitem"),10)
	mwgubun     = requestCheckvar(request("mwgubun"),10)
	vatyn       = requestCheckvar(request("vatyn"),10)
	minusinc   = requestCheckvar(request("minusinc"),10)
	bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
	buseo       = requestCheckvar(request("buseo"),10)
	itemgubun   = requestCheckvar(request("itemgubun"),10)
	purchasetype   = requestCheckvar(request("purchasetype"),10)
	stplace     = requestCheckvar(request("stplace"),10)
	showsuply   = requestCheckvar(request("showsuply"),10)
	dtype       = requestCheckvar(request("dtype"),10)
	makerid     = requestCheckvar(request("makerid"),32)
	shopid     = requestCheckvar(request("shopid"),32)
	etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)
	showDiff      = requestCheckvar(request("showDiff"),10)
	brandUseYN      = requestCheckvar(request("brandUseYN"),10)

if (makerid<>"") then dtype=""
if (sysorreal="") then sysorreal="sys"  ''real
if (research="") and (bPriceGbn = "") then
    bPriceGbn="V"
end if
if (stplace="") then
    stplace="L"
	showDiff = "Y"
end if
if (research="") then
	if (itemgubun = "") then
		'itemgubun = "AA"
	end if
	if (buseo = "") then
		buseo = "3X"
	end if
end if

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
	ojaego.FRectPurchaseType = purchasetype
	ojaego.FRectShopSuplyPrice = CHKIIF(showsuply="on",1,0)
	ojaego.FRectPlaceGubun = stplace
	ojaego.FRectShopID    = shopid
	ojaego.FRectetcjungsantype = etcjungsantype

	if (dtype="mk") then
	    ojaego.FRectGroupbyType = CHKIIF(dtype="mk",1,0)
	end if

	ojaego.FRectMakerid = makerid
	'ojaego.FRectIsUsing = isusing
	'ojaego.FRectGubun = sysorreal
	'ojaego.FRectNewItem = newitem

	''ojaego.FRectMinusInclude = minusinc

	'if (buseo="IT") then
	'    ojaego.FRectITSOnlyOrNot = "O"
	'else
	'    ojaego.FRectITSOnlyOrNot = "N"
	'end if
	'
	ojaego.FRectPriceGubun = bPriceGbn
	ojaego.FRectBrandUseYN = brandUseYN
	ojaego.GetMonthlyJeagoSumSummary '' GetMonthlyJeagoSumWithPreMonth ''

dim totno, totbuy, subTotno, subTotbuy, totPreno, totPrebuy , subPreno, subPrebuy, totIpno,totIpBuy , subIpno, subIpBuy ', totavgBuy, offtotavgBuy
dim totLossno, totLossBuy , subLossno, subLossBuy, totSellno, totSellBuy , subSellno, subSellBuy, isItemList, isGroupByBrand
dim totOffChulno, totOffChulBuy , subOffChulno, subOffChulBuy, totEtcChulno, totEtcChulBuy , subEtcChulno, subEtcChulBuy
dim totCsChulno, totCsChulBuy  , subCsChulno, subCsChulBuy, iURL, iURLEtc, nBusiName, diffStock, diffStockPrc, diffStockW
dim totErrBadItemno, totErrBadItemBuy, subErrBadItemno, subErrBadItemBuy, totMoveItemno, totMoveItemBuy, subMoveItemno, subMoveItemBuy
dim totErrRealCheckno, totErrRealCheckBuy, subErrRealCheckno, subErrRealCheckBuy, totErrRealCheckBuyPlus, totErrRealCheckBuyMinus
dim totRealStockno, totRealStockBuy, subRealStockno, subRealStockBuy
dim totPreRealStockno, totPreRealStockbuy, subPreRealStockno, subPreRealStockbuy

isGroupByBrand = (dtype="mk")
isItemList = (makerid<>"")
%>
<script type="text/javascript">

function goSubmit(imakerid){
	frm.target='';
	frm.action="";
	frm.submit();
}

function fnResearchByBrand(imakerid){
    document.frm.makerid.value=imakerid;
    document.frm.target="_blank";
	frm.action="";
    document.frm.submit();
}

// 누적재고
function reActAccMonthSummary(){
    alert("수정중");
    return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + " 월말 재고 내역을 재작성 하시겠습니까?")){ return; }

	var popwin = window.open("do_stocksummary.asp?mode=monthlystock&yyyymm=" + yyyymm,"reActAccMonthSummary","width=100,height=100");
	popwin.focus();
}

function reActAccMonthSummaryEXL(){
	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + " 월말 재고 내역(엑셀)을 재작성 하시겠습니까?")){ return; }

	//var popwin = window.open("do_stocksummary.asp?mode=monthlystockexl&yyyymm=" + yyyymm,"reActAccMonthSummaryEXL","width=100,height=100");
	//popwin.focus();

	window.open('','makerFileconfirm','width=1400,height=800,scrollbars=yes,resizable=yes');
	frm.target='makerFileconfirm';
	frm.action="/admin/newreport/do_stocksummary.asp";
	frm.mode.value="monthlystockexl";
	frm.submit();
	makerFileconfirm.focus();
	frm.mode.value="";
	frm.target='';
	frm.action="";
}

// 누적재고+서머리정보(위탁상품)
function reActAccMonthSummaryOneItem(itemgubun, itemid, itemoption){
	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + " 월말 재고 내역(위탁상품)을 재작성 하시겠습니까?")){ return; }

	var popwin = window.open("do_stocksummary.asp?mode=monthlystockoneitem&yyyymm=" + yyyymm + "&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"reActAccMonthSummaryOneItem","width=100,height=100");
	popwin.focus();
}

// 서머리정보
function reActMonthSummary(){
	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + " 월말 재고 내역을 재작성 하시겠습니까?")){ return; }

	var popwin = window.open("do_stocksummary.asp?mode=monthlystocksum&yyyymm=" + yyyymm,"reActMonthSummary","width=100,height=100");
	popwin.focus();
}

//재고 pop
function TnPopItemStockWithGubun(itemgubun,itemid,itemoption,shopid){
	<% if (stplace = "S") then %>
	var popwin = window.open("/common/offshop/shop_itemcurrentstock.asp?shopid="+shopid+"&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"jsSearchItemStock","width=1000 height=600 scrollbars=yes resizable=yes");
	<% else %>
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"jsSearchItemStock","width=1200 height=600 scrollbars=yes resizable=yes");
	<% end if %>

	popwin.focus();
}

function popAccStockModiOne(itemgubun,itemid,itemoption){
    <% if (stplace = "S") and (shopid = "") then %>
	alert("먼저 매장을 선택하세요.");
	return;
	<% end if %>
	var popwin = window.open("/admin/newreport/pop_item_stock_Accsummary_edit.asp?yyyy1=<%=yyyy1%>&mm1=<%=mm1%>&shopid=<%= shopid %>&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"popAccStockModiOne","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popXL() {
	//var popwin = window.open("monthlystock_xl_download.asp?yyyymm=<%'= (yyyy1 + "-" + mm1) %>&placeGubun=<%'= stplace %>&priceGubun=<%'= bPriceGbn %>","reActAccMonthSummary","width=1000,height=1000 scrollbars=yes resizable=yes");
	//popwin.focus();

	window.open('','downFileconfirm','width=300,height=300,scrollbars=yes,resizable=yes');
	frm.target='downFileconfirm';
	frm.action="/admin/newreport/monthlystock_xl_download.asp";
	frm.submit();
	downFileconfirm.focus();
	frm.target='';
	frm.action="";
}

function popRealStockXL() {
	window.open('','downFileconfirmRealStock','width=300,height=300,scrollbars=yes,resizable=yes');
	frm.target='downFileconfirmRealStock';
	frm.action="/admin/newreport/monthlystock_xl_download_realstock.asp";
	frm.submit();
	downFileconfirmRealStock.focus();
	frm.target='';
	frm.action="";
}

</script>
<!-- 검색 시작 -->
<form name="frm" method="get" action="" target="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 말일자 재고자산
        &nbsp;&nbsp;|&nbsp;&nbsp;
    	과세구분
    	<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >전체
    	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >과세
    	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >면세
    	&nbsp;&nbsp;<input type="checkbox" name="showsuply" value="on" <%= CHKIIF(showsuply="on","checked","") %> >공급가로 표시

        <% if makerid<>"" then %>
	        &nbsp;&nbsp;|&nbsp;&nbsp;
	        브랜드 : <% drawSelectBoxDesignerwithName "makerid",makerid %>
        <% else %>
        	<input type="hidden" name="makerid" value="">
	    	<% if (dtype<>"") then %>
	    		&nbsp;&nbsp;|&nbsp;&nbsp;
	    		그루핑
	    		<input type="radio" name="dtype" value="" <% if dtype="" then response.write "checked" %>> 상품구분
	    		<input type="radio" name="dtype" value="mk" <% if dtype="mk" then response.write "checked" %>> 브랜드
	    	<% end if %>
    	<% end if %>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="goSubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<font color="#CC3333">재고구분:</font>
    	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >시스템재고
    	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >실사재고
    	&nbsp;&nbsp;
	    <% if (FALSE)  then %>
	    	<font color="#CC3333">상품사용구분:</font>
	    	<input type="radio" name="isusing" value="" <% if isusing="" then response.write "checked" %> >전체
	    	<input type="radio" name="isusing" value="Y" <% if isusing="Y" then response.write "checked" %> >사용함
	    	<input type="radio" name="isusing" value="N" <% if isusing="N" then response.write "checked" %> >사용안함
	    	&nbsp;&nbsp;
	    <% end if %>
    	<font color="#CC3333">매입구분:</font>
    	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >전체
    	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >매입
    	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >위탁
    	<!-- <input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >업체 -->
    	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >미지정
        <% if (mwgubun<>"" and mwgubun<>"M" and mwgubun<>"W" and mwgubun<>"Z") then %>
            <input type="radio" name="mwgubun" value="<%=mwgubun%>" checked ><%=mwgubun%>
		<% end if %>
		<font color="#CC3333">브랜드사용구분:</font>
		<select class="select" name="brandUseYN">
			<option value=""></option>
			<option value="Y" <%= CHKIIF(brandUseYN="Y", "selected", "") %> >사용</option>
			<option value="N" <%= CHKIIF(brandUseYN="N", "selected", "") %> >사용안함</option>
		</select>

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
	    <input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %> disabled >평균매입가
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
        <% Call drawSelectBoxBuseoGubunWith3PL("buseo", buseo) %>
		&nbsp;&nbsp;&nbsp;
    	<font color="#CC3333">상품구분:</font>
		<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
		&nbsp;&nbsp;&nbsp;
		구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
		<% if (stplace = "S") then %>
			&nbsp;
			매장(매장재고 검색시) : <% Call drawSelectBoxAccShop(yyyy1 + "-" + mm1, "", "shopid", shopid) %>

			&nbsp;
			정산방법:
			<select class="select" name="etcjungsantype"  >
            <option value="">-선택-</option>
            <option value="1" <%=CHKIIF(etcjungsantype="1","selected","")%> >판매분정산</option>
            <option value="2" <%=CHKIIF(etcjungsantype="2","selected","")%> >출고분정산</option>
            <option value="3" <%=CHKIIF(etcjungsantype="3","selected","")%> >가맹점정산</option>
            <option value="4" <%=CHKIIF(etcjungsantype="4","selected","")%> >직영점정산</option>
            <option value="41" <%=CHKIIF(etcjungsantype="41","selected","")%> >직영점+판매분정산</option>
            </select>
		<% end if %>
		&nbsp;&nbsp;&nbsp;<input type="checkbox" name="showDiff" value="Y" <%= CHKIIF(showDiff="Y","checked","") %> > 오차표시
    </td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

* TODO : 아울렛 재고 평균매입가는 원상품의 평균매입가가 입력되어야 한다.<br />
* 전월재고수량이 "<font color="red">재고자산(월별)</font>" 과 일치하지 않으면 전월재고를 재작성(서머리) 하면 된다.<br />
* 오늘 작성한 입출고내역은 "<font color="red">누적재고</font>" 를 재작성해야 반영됩니다.
<% if (C_ADMIN_AUTH) then %>
	<!-- 액션 시작 -->
	<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="재작성(서머리)" onclick="reActMonthSummary();">
			<input type="button" class="button" value="재작성(누적재고)" onclick="reActAccMonthSummary();">
		</td>
		<td align="right">
			<input type="button" class="button" value="엑셀자료생성" onclick="reActAccMonthSummaryEXL();">
            <input type="button" class="button" value="엑셀받기(<%=CHKIIF(bPriceGbn="P","작성시매입가","평균매입가")%>,시스템재고)" onclick="popXL();">
			<input type="button" class="button" value="엑셀받기(<%=CHKIIF(bPriceGbn="P","작성시매입가","평균매입가")%>,실사재고)" onclick="popRealStockXL();">
		</td>
	</tr>
	</table>
	<!-- 액션 끝 -->
<% end if %>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if (isGroupByBrand) then %>
		<td colspan="2">브랜드</td>
    <% else %>
		<td colspan="6">상품구분</td>
    <% end if %>
	<td colspan="2">
		<% if sysorreal="real" then %>
			기초재고(실사재고)
		<% else %>
			기초재고(시스템)
		<% end if %>
		<br>A
	</td>
    <td colspan="2">당월매입(월)<br>B</td>
	<td colspan="2">당월이동(월)<br>M</td>
    <td colspan="2">당월판매(월)<br>S</td>
    <td colspan="2">당월출고1(월)<br>C</td>
    <td colspan="2">당월출고2(월)<br>E1</td>
    <td colspan="2">당월기타출고(월)<br>E2</td>
    <td colspan="2">당월CS출고(월)<br>E</td>
	<td colspan="2">오차(월)</td>
    <td colspan="2">
		<% if sysorreal="real" then %>
			기말재고(실사재고)<br>L
		<% else %>
			기말재고(시스템)<br>L
		<% end if %>
	</td>
	<td colspan="2">누적오차</td>
	<% if sysorreal="sys" then %>
		<td colspan="2">기말재고(실사)</td>
	<% end if %>
	<td colspan="2">누적불량</td>
    <% if (isViewWonga) then %>
	    <td colspan="2">총매출원가<br>D=A+B-L</td>
	    <td width="1" bgcolor="#FFFFFF"></td>
	    <td colspan="2">LOSS출고<br>E</td>
	    <td colspan="2">상품매출원가<br>F=A+B+E-L</td>
    <% end if %>
    <% if (isItemList) then %>
		<td rowspan="2">비고<br>검토</td>
    <% end if %>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <% if (isGroupByBrand) then %>
		<td >브랜드ID</td>
		<td >구매유형</td>
    <% else %>
	    <td width="60">부서</td>
	    <td width="60">구분</td>
	    <td width="30">코드<% if makerid="" then%><br>구분<% end if %></td>
	    <td>상품명</td>
		<td width="30">매입<br>구분</td>
		<td >
			<% if (makerid = "") then %>
				매장
			<% else %>
				최종입고
			<% end if %>
		</td>
	<% end if %>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<% if sysorreal="sys" then %>
		<td>수량</td>
		<td>금액<br>(매입가)</td>
	<% end if %>
	<td>수량</td>
	<td>금액<br>(매입가)</td>
	<% if (isViewWonga) then %>
		<td width="50">수량</td>
		<td width="80">금액<br>(매입가)</td>
		<td  bgcolor="#FFFFFF"></td>
		<td width="50">수량</td>
		<td width="80">금액<br>(매입가)</td>
		<td width="50">수량</td>
		<td width="80">금액<br>(매입가)</td>
	<% end if %>
</tr>
<% for i=0 to ojaego.FResultCount-1 %>
<%
if (ojaego.FItemList(i).Fitemgubun <> "75") and (ojaego.FItemList(i).Fitemgubun <> "80") and (ojaego.FItemList(i).Fitemgubun <> "85") then

totno       = totno + ojaego.FItemList(i).FTotCount
totbuy      = totbuy + ojaego.FItemList(i).FTotBuySum
totPreno    = totPreno + ojaego.FItemList(i).FTotPreCount
totPrebuy   = totPrebuy + ojaego.FItemList(i).FTotPreBuySum
totPreRealStockno    = totPreRealStockno + ojaego.FItemList(i).FTotPreRealStockCount
totPreRealStockbuy   = totPreRealStockbuy + ojaego.FItemList(i).FTotPreRealStockBuySum
totIpno     = totIpno + ojaego.FItemList(i).FTotIpCount
totIpBuy    = totIpBuy + ojaego.FItemList(i).FTotIpBuySum

totLossno   = totLossno + ojaego.FItemList(i).FTotLossCount
totLossBuy  = totLossBuy + ojaego.FItemList(i).FTotLossBuySum

totSellno       = totSellno + ojaego.FItemList(i).FTotSellCount
totSellBuy      = totSellBuy + ojaego.FItemList(i).FTotSellBuySum
totOffChulno    = totOffChulno + ojaego.FItemList(i).FTotOffChulCount
totOffChulBuy   = totOffChulBuy + ojaego.FItemList(i).FTotOffChulBuySum
totEtcChulno    = totEtcChulno + ojaego.FItemList(i).FTotEtcChulCount
totEtcChulBuy   = totEtcChulBuy + ojaego.FItemList(i).FTotEtcChulBuySum
totCsChulno     = totCsChulno + ojaego.FItemList(i).FTotCsChulCount
totCsChulBuy    = totCsChulBuy + ojaego.FItemList(i).FTotCsChulBuySum
subTotno        = subTotno + ojaego.FItemList(i).FTotCount
subTotbuy       = subTotbuy + ojaego.FItemList(i).FTotBuySum
subPreno        = subPreno + ojaego.FItemList(i).FTotPreCount
subPrebuy       = subPrebuy + ojaego.FItemList(i).FTotPreBuySum
subPreRealStockno        = subPreRealStockno + ojaego.FItemList(i).FTotPreRealStockCount
subPreRealStockbuy       = subPreRealStockbuy + ojaego.FItemList(i).FTotPreRealStockBuySum
subIpno         = subIpno + ojaego.FItemList(i).FTotIpCount
subIpBuy        = subIpBuy + ojaego.FItemList(i).FTotIpBuySum
subLossno       = subLossno + ojaego.FItemList(i).FTotLossCount
subLossBuy      = subLossBuy + ojaego.FItemList(i).FTotLossBuySum

subSellno       = subSellno + ojaego.FItemList(i).FTotSellCount
subSellBuy      = subSellBuy + ojaego.FItemList(i).FTotSellBuySum
subOffChulno    = subOffChulno + ojaego.FItemList(i).FTotOffChulCount
subOffChulBuy   = subOffChulBuy + ojaego.FItemList(i).FTotOffChulBuySum
subEtcChulno    = subEtcChulno + ojaego.FItemList(i).FTotEtcChulCount
subEtcChulBuy   = subEtcChulBuy + ojaego.FItemList(i).FTotEtcChulBuySum
subCsChulno     = subCsChulno + ojaego.FItemList(i).FTotCsChulCount
subCsChulBuy    = subCsChulBuy + ojaego.FItemList(i).FTotCsChulBuySum


totErrBadItemno     = totErrBadItemno + ojaego.FItemList(i).FTotErrBadItemCount
totErrBadItemBuy    = totErrBadItemBuy + ojaego.FItemList(i).FTotErrBadItemBuySum
subErrBadItemno     = subErrBadItemno + ojaego.FItemList(i).FTotErrBadItemCount
subErrBadItemBuy    = subErrBadItemBuy + ojaego.FItemList(i).FTotErrBadItemBuySum

totErrRealCheckno     = totErrRealCheckno + ojaego.FItemList(i).FTotErrRealCheckCount
totErrRealCheckBuy    = totErrRealCheckBuy + ojaego.FItemList(i).FTotErrRealCheckBuySum
subErrRealCheckno     = subErrRealCheckno + ojaego.FItemList(i).FTotErrRealCheckCount
subErrRealCheckBuy    = subErrRealCheckBuy + ojaego.FItemList(i).FTotErrRealCheckBuySum

totRealStockno     = totRealStockno + ojaego.FItemList(i).FTotRealStockCount
totRealStockBuy    = totRealStockBuy + ojaego.FItemList(i).FTotRealStockBuySum
subRealStockno     = subRealStockno + ojaego.FItemList(i).FTotRealStockCount
subRealStockBuy    = subRealStockBuy + ojaego.FItemList(i).FTotRealStockBuySum

totMoveItemno     = totMoveItemno + ojaego.FItemList(i).FTotMoveItemCount
totMoveItemBuy    = totMoveItemBuy + ojaego.FItemList(i).FTotMoveItemBuySum
subMoveItemno     = subMoveItemno + ojaego.FItemList(i).FTotMoveItemCount
subMoveItemBuy    = subMoveItemBuy + ojaego.FItemList(i).FTotMoveItemBuySum

if (ojaego.FItemList(i).FTotErrRealCheckBuySum > 0) then
	totErrRealCheckBuyPlus = totErrRealCheckBuyPlus + ojaego.FItemList(i).FTotErrRealCheckBuySum
else
	totErrRealCheckBuyMinus = totErrRealCheckBuyMinus + ojaego.FItemList(i).FTotErrRealCheckBuySum
end if


iURL = "monthlystock_summary.asp?menupos="& menupos &"&dtype=mk&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype &"&stplace="&stplace &"&shopid="&ojaego.FItemList(i).Fshopid&"&etcjungsantype="&etcjungsantype & "&showDiff="&showDiff
if Not(isOnlySys) THEN iURL=iURL&"&sysorreal="& sysorreal

iURLEtc = "monthlystock_etcChulgoList.asp?menupos="& menupos &"&dtype=mk&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
iURLEtc = iURLEtc + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype &"&stplace="&stplace &"&shopid="&shopid&"&etcjungsantype="&etcjungsantype
if Not(isOnlySys) THEN iURLEtc=iURLEtc&"&sysorreal="& sysorreal
%>
<% if sysorreal="real" then %>
	<tr align="right" bgcolor="<%=CHKIIF((isItemList or isGroupByBrand) and ojaego.FItemList(i).getCalcuCurRealStock<>ojaego.FItemList(i).FTotRealStockCount,"yellow","#FFFFFF")%>" >
<% else %>
	<tr align="right" bgcolor="<%=CHKIIF((isItemList or isGroupByBrand) and ojaego.FItemList(i).getCalcuCurSysStock<>ojaego.FItemList(i).FTotCount,"yellow","#FFFFFF")%>" >
<% end if %>
    <% if (isGroupByBrand) then %>
	    <td align="center">
			<a href="javascript:fnResearchByBrand('<% if (ojaego.FItemList(i).Fmakerid <> "") then %><%= ojaego.FItemList(i).Fmakerid %><% else %>-<% end if %>');">
				<% if (ojaego.FItemList(i).Fmakerid <> "") then %><%= ojaego.FItemList(i).Fmakerid %><% else %>-<% end if %>
			</a>
		</td>
		<td align="center"><%= ojaego.FItemList(i).fpurchasetypename %></td>
    <% else %>
	    <td align="center">
	        <% if makerid<>"" then%>
				<%= ojaego.FItemList(i).Fshopid %>
	        <% else %>
				<%= ojaego.FItemList(i).getBusiName %>
	        <% end if %>
	    </td>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= GetItemGubunName(ojaego.FItemList(i).Fitemgubun) %></a></td>
		<td align="center">
		    <% if makerid<>"" then%>
				<a href="javascript:TnPopItemStockWithGubun('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>','<%= ojaego.FItemList(i).Fshopid %>')">
				<%= ojaego.FItemList(i).getLogisticsCode%></a>
		    <% else %>
				<a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).Fitemgubun %></a>
		    <% end if %>
		</td>
		<td align="left">
		    <%= ojaego.FItemList(i).fitemname %>
		</td>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
		<td align="center">
			<% if (makerid = "") then %>
			<%= ojaego.FItemList(i).Fshopid %>
			<% else %>
			<%= ojaego.FItemList(i).FlastIpgoDate %>
			<% end if %>
		</td>
	<% end if %>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber(ojaego.FItemList(i).FTotPreRealStockCount,0) %>
		<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FTotPreCount,0) %>
		<% end if %>
	</td>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber(ojaego.FItemList(i).FTotPreRealStockBuySum,0) %>
		<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FTotPreBuySum,0) %>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotIpCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotIpBuySum,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotMoveItemCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotMoveItemBuySum,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotSellCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotSellBuySum,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotOffChulCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotOffChulBuySum,0) %></td>

	<% if makerid<>"" then%>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotEtcChulCount,0) %></td>
	    <td><%= FormatNumber(ojaego.FItemList(i).FTotEtcChulBuySum,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></td>
	<% else %>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotEtcChulCount,0) %></td>
	    <td><%= FormatNumber(ojaego.FItemList(i).FTotEtcChulBuySum,0) %></td>
		<td><a href="<%= iURLEtc %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></a></td>
		<td><a href="<%= iURLEtc %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></a></td>
	<% end if %>

	<td><%= FormatNumber(ojaego.FItemList(i).FTotCsChulCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotCsChulBuySum,0) %></td>
	<%
	if sysorreal="real" then
		diffStock=diffStock+(ojaego.FItemList(i).getCalcuCurRealStock-ojaego.FItemList(i).FTotRealStockCount)*-1
		diffStockPrc=diffStockPrc+(ojaego.FItemList(i).getCalcuCurRealBuySum-ojaego.FItemList(i).FTotRealStockBuySum)*-1
	else
		diffStock=diffStock+(ojaego.FItemList(i).getCalcuCurSysStock-ojaego.FItemList(i).FTotCount)*-1
		diffStockPrc=diffStockPrc+(ojaego.FItemList(i).getCalcuCurSysBuySum-ojaego.FItemList(i).FTotBuySum)*-1
	end if
	%>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber((ojaego.FItemList(i).getCalcuCurRealStock-ojaego.FItemList(i).FTotRealStockCount)*-1,0) %>
		<% else %>
			<%= FormatNumber((ojaego.FItemList(i).getCalcuCurSysStock-ojaego.FItemList(i).FTotCount)*-1,0) %>
		<% end if %>
	</td>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber((ojaego.FItemList(i).getCalcuCurRealBuySum-ojaego.FItemList(i).FTotRealStockBuySum)*-1,0) %>
		<% else %>
			<%= FormatNumber((ojaego.FItemList(i).getCalcuCurSysBuySum-ojaego.FItemList(i).FTotBuySum)*-1,0) %>
		<% end if %>
	</td>
    <td>
		<% if sysorreal="real" then %>
			<%= FormatNumber(ojaego.FItemList(i).FTotRealStockCount,0) %>
		<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %>
		<% end if %>
    </td>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber(ojaego.FItemList(i).FTotRealStockBuySum,0) %>
		<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotErrRealCheckCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotErrRealCheckBuySum,0) %></td>
	<% if sysorreal="sys" then %>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotRealStockCount,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotRealStockBuySum,0) %></td>
	<% end if %>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotErrBadItemCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotErrBadItemBuySum,0) %></td>
	<% if (isViewWonga) then %>
		<td>
			<%= FormatNumber(ojaego.FItemList(i).getWongaCnt,0) %>

			<% if sysorreal="real" then %>
				<% if (ojaego.FItemList(i).getCalcuCurRealStock<>ojaego.FItemList(i).FTotPreRealStockCount) and (showDiff = "Y") then %>
					<% diffStockW=diffStockW+ojaego.FItemList(i).FTotPreRealStockCount-ojaego.FItemList(i).getCalcuCurRealStock %>
					<br><font color="red"><%= FormatNumber(ojaego.FItemList(i).FTotPreRealStockCount-ojaego.FItemList(i).getCalcuCurRealStock,0) %></font>
				<% end if %>
			<% else %>
				<% if (ojaego.FItemList(i).getCalcuCurSysStock<>ojaego.FItemList(i).FTotCount) and (showDiff = "Y") then %>
					<% diffStockW=diffStockW+ojaego.FItemList(i).FTotCount-ojaego.FItemList(i).getCalcuCurSysStock %>
					<br><font color="red"><%= FormatNumber(ojaego.FItemList(i).FTotCount-ojaego.FItemList(i).getCalcuCurSysStock,0) %></font>
				<% end if %>
			<% end if %>
		</td>
		<td><%= FormatNumber(ojaego.FItemList(i).getWongaSum,0) %></td>
		<td align="center" valign="middle">
			<% if makerid<>"" then %>
				<% if sysorreal="real" then %>
					<% if (isItemList or isGroupByBrand) and ojaego.FItemList(i).getCalcuCurRealStock<>ojaego.FItemList(i).FTotPreRealStockCount then %>
						<a href="javascript:reActAccMonthSummaryOneItem('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>')"><img src="/images/icon_reload.gif" border="0"></a>
					<% end if %>
				<% else %>
					<% if (isItemList or isGroupByBrand) and ojaego.FItemList(i).getCalcuCurSysStock<>ojaego.FItemList(i).FTotCount then %>
						<a href="javascript:reActAccMonthSummaryOneItem('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>')"><img src="/images/icon_reload.gif" border="0"></a>
					<% end if %>
				<% end if %>
			<% end if %>
		</td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaCnt,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaSum,0) %></td>
	<% end if %>
	<% if (isItemList) then %>
		<td align="center"><img src="/images/icon_arrow_link.gif" style="cursor:pointer" onClick="popAccStockModiOne('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>')"></td>
    <% end if %>
</tr>
<%
end if
next
%>
<% if (Not isGroupByBrand) then %>
	<tr align="right" bgcolor="#EEFFEE">
		<td></td>
		<td align="center">상품소계</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td>
			<% if sysorreal="real" then %>
				<%= FormatNumber(subPreRealStockno,0) %>
			<% else %>
				<%= FormatNumber(subPreno,0) %>
			<% end if %>
		</td>
		<td>
			<% if sysorreal="real" then %>
				<%= FormatNumber(subPreRealStockbuy,0) %>
			<% else %>
				<%= FormatNumber(subPrebuy,0) %>
			<% end if %>
		</td>
		<td><%= FormatNumber(subIpno,0) %></td>
		<td><%= FormatNumber(subIpBuy,0) %></td>
		<td><%= FormatNumber(subMoveItemno,0) %></td>
		<td><%= FormatNumber(subMoveItemBuy,0) %></td>
		<td><%= FormatNumber(subSellno,0) %></td>
		<td><%= FormatNumber(subSellBuy,0) %></td>
		<td><%= FormatNumber(subOffChulno,0) %></td>
		<td><%= FormatNumber(subOffChulBuy,0) %></td>
		<td><%= FormatNumber(subEtcChulno,0) %></td>
		<td><%= FormatNumber(subEtcChulBuy,0) %></td>
		<td><%= FormatNumber(subLossno,0) %></td>
		<td><%= FormatNumber(subLossBuy,0) %></td>
		<td><%= FormatNumber(subCsChulno,0) %></td>
		<td><%= FormatNumber(subCsChulBuy,0) %></td>
		<td><%= FormatNumber(diffStock,0) %></td>
		<td><%= FormatNumber(diffStockPrc,0) %></td>
		<td>
			<% if sysorreal="real" then %>
				<%= FormatNumber(subRealStockno,0) %>
			<% else %>
				<%= FormatNumber(subTotno,0) %>
			<% end if %>
		</td>
		<td>
			<% if sysorreal="real" then %>
				<%= FormatNumber(subRealStockBuy,0) %>
			<% else %>
				<%= FormatNumber(subTotbuy,0) %>
			<% end if %>
		</td>
		<td><%= FormatNumber(subErrRealCheckno,0) %></td>
		<td><%= FormatNumber(subErrRealCheckBuy,0) %></td>
		<% if sysorreal="sys" then %>
			<td><%= FormatNumber(subRealStockno,0) %></td>
			<td><%= FormatNumber(subRealStockBuy,0) %></td>
		<% end if %>
		<td><%= FormatNumber(subErrBadItemno,0) %></td>
		<td><%= FormatNumber(subErrBadItemBuy,0) %></td>
		<% if (isViewWonga) then %>
			<td>
				<% if sysorreal="real" then %>
					<%= FormatNumber(subPreRealStockno+subIpno-subRealStockno,0) %>
				<% else %>
					<%= FormatNumber(subPreno+subIpno-subtotno,0) %>
				<% end if %>
			</td>
			<td>
				<% if sysorreal="real" then %>
					<%= FormatNumber(subPreRealStockbuy+subIpBuy-subRealStockBuy,0) %>
				<% else %>
					<%= FormatNumber(subPrebuy+subIpBuy-subtotbuy,0) %>
				<% end if %>
			</td>
			<td></td>
			<td><%= FormatNumber(subLossno,0) %></td>
			<td><%= FormatNumber(subLossBuy,0) %></td>
			<td>
				<% if sysorreal="real" then %>
					<%= FormatNumber(subPreRealStockno+subIpno-subRealStockno+subLossno,0) %>
				<% else %>
					<%= FormatNumber(subPreno+subIpno-subtotno+subLossno,0) %>
				<% end if %>
			</td>
			<td>
				<% if sysorreal="real" then %>
					<%= FormatNumber(subPreRealStockbuy+subIpBuy-subRealStockBuy+subLossBuy,0) %>
				<% else %>
					<%= FormatNumber(subPrebuy+subIpBuy-subtotbuy+subLossBuy,0) %>
				<% end if %>
			</td>
		<% end if %>
		<% if (isItemList) then %>
			<td></td>
		<% end if %>
	</tr>
	<!--<tr  bgcolor="#FFFFFF"><td colspan="32"></td></tr>-->
	<%
	subTotno=0
	subTotbuy=0
	subPreno   =0
	subPrebuy  =0
	subPreRealStockno   =0
	subPreRealStockbuy  =0
	subIpno    =0
	subIpBuy    =0
	subLossno   =0
	subLossBuy  =0

	subSellno       = 0
	subSellBuy      = 0
	subOffChulno    = 0
	subOffChulBuy   = 0
	subEtcChulno    = 0
	subEtcChulBuy   = 0
	subCsChulno     = 0
	subCsChulBuy    = 0
	subErrBadItemno = 0
	subErrBadItemBuy = 0
	subErrRealCheckno = 0
	subErrRealCheckBuy = 0
	subRealStockno = 0
	subRealStockBuy = 0
	subMoveItemno = 0
	subMoveItemBuy = 0
	%>
<% end if %>
<% for i=0 to ojaego.FResultCount-1 %>
<%
if (ojaego.FItemList(i).Fitemgubun = "75") or (ojaego.FItemList(i).Fitemgubun = "80") or (ojaego.FItemList(i).Fitemgubun = "85") then

totno       = totno + ojaego.FItemList(i).FTotCount
totbuy      = totbuy + ojaego.FItemList(i).FTotBuySum
totPreno    = totPreno + ojaego.FItemList(i).FTotPreCount
totPrebuy   = totPrebuy + ojaego.FItemList(i).FTotPreBuySum
totPreRealStockno    = totPreRealStockno + ojaego.FItemList(i).FTotPreRealStockCount
totPreRealStockbuy   = totPreRealStockbuy + ojaego.FItemList(i).FTotPreRealStockBuySum
totIpno     = totIpno + ojaego.FItemList(i).FTotIpCount
totIpBuy    = totIpBuy + ojaego.FItemList(i).FTotIpBuySum

totLossno   = totLossno + ojaego.FItemList(i).FTotLossCount
totLossBuy  = totLossBuy + ojaego.FItemList(i).FTotLossBuySum

totSellno       = totSellno + ojaego.FItemList(i).FTotSellCount
totSellBuy      = totSellBuy + ojaego.FItemList(i).FTotSellBuySum
totOffChulno    = totOffChulno + ojaego.FItemList(i).FTotOffChulCount
totOffChulBuy   = totOffChulBuy + ojaego.FItemList(i).FTotOffChulBuySum
totEtcChulno    = totEtcChulno + ojaego.FItemList(i).FTotEtcChulCount
totEtcChulBuy   = totEtcChulBuy + ojaego.FItemList(i).FTotEtcChulBuySum
totCsChulno     = totCsChulno + ojaego.FItemList(i).FTotCsChulCount
totCsChulBuy    = totCsChulBuy + ojaego.FItemList(i).FTotCsChulBuySum

subTotno        = subTotno + ojaego.FItemList(i).FTotCount
subTotbuy       = subTotbuy + ojaego.FItemList(i).FTotBuySum
subPreno        = subPreno + ojaego.FItemList(i).FTotPreCount
subPrebuy       = subPrebuy + ojaego.FItemList(i).FTotPreBuySum
subPreRealStockno        = subPreRealStockno + ojaego.FItemList(i).FTotPreRealStockCount
subPreRealStockbuy       = subPreRealStockbuy + ojaego.FItemList(i).FTotPreRealStockBuySum
subIpno         = subIpno + ojaego.FItemList(i).FTotIpCount
subIpBuy        = subIpBuy + ojaego.FItemList(i).FTotIpBuySum
subLossno       = subLossno + ojaego.FItemList(i).FTotLossCount
subLossBuy      = subLossBuy + ojaego.FItemList(i).FTotLossBuySum

subSellno       = subSellno + ojaego.FItemList(i).FTotSellCount
subSellBuy      = subSellBuy + ojaego.FItemList(i).FTotSellBuySum
subOffChulno    = subOffChulno + ojaego.FItemList(i).FTotOffChulCount
subOffChulBuy   = subOffChulBuy + ojaego.FItemList(i).FTotOffChulBuySum
subEtcChulno    = subEtcChulno + ojaego.FItemList(i).FTotEtcChulCount
subEtcChulBuy   = subEtcChulBuy + ojaego.FItemList(i).FTotEtcChulBuySum
subCsChulno     = subCsChulno + ojaego.FItemList(i).FTotCsChulCount
subCsChulBuy    = subCsChulBuy + ojaego.FItemList(i).FTotCsChulBuySum


totErrBadItemno     = totErrBadItemno + ojaego.FItemList(i).FTotErrBadItemCount
totErrBadItemBuy    = totErrBadItemBuy + ojaego.FItemList(i).FTotErrBadItemBuySum
subErrBadItemno     = subErrBadItemno + ojaego.FItemList(i).FTotErrBadItemCount
subErrBadItemBuy    = subErrBadItemBuy + ojaego.FItemList(i).FTotErrBadItemBuySum

totErrRealCheckno     = totErrRealCheckno + ojaego.FItemList(i).FTotErrRealCheckCount
totErrRealCheckBuy    = totErrRealCheckBuy + ojaego.FItemList(i).FTotErrRealCheckBuySum
subErrRealCheckno     = subErrRealCheckno + ojaego.FItemList(i).FTotErrRealCheckCount
subErrRealCheckBuy    = subErrRealCheckBuy + ojaego.FItemList(i).FTotErrRealCheckBuySum

totRealStockno     = totRealStockno + ojaego.FItemList(i).FTotRealStockCount
totRealStockBuy    = totRealStockBuy + ojaego.FItemList(i).FTotRealStockBuySum
subRealStockno     = subRealStockno + ojaego.FItemList(i).FTotRealStockCount
subRealStockBuy    = subRealStockBuy + ojaego.FItemList(i).FTotRealStockBuySum

totMoveItemno     = totMoveItemno + ojaego.FItemList(i).FTotMoveItemCount
totMoveItemBuy    = totMoveItemBuy + ojaego.FItemList(i).FTotMoveItemBuySum
subMoveItemno     = subMoveItemno + ojaego.FItemList(i).FTotMoveItemCount
subMoveItemBuy    = subMoveItemBuy + ojaego.FItemList(i).FTotMoveItemBuySum

if (ojaego.FItemList(i).FTotErrRealCheckBuySum > 0) then
	totErrRealCheckBuyPlus = totErrRealCheckBuyPlus + ojaego.FItemList(i).FTotErrRealCheckBuySum
else
	totErrRealCheckBuyMinus = totErrRealCheckBuyMinus + ojaego.FItemList(i).FTotErrRealCheckBuySum
end if


iURL = "monthlystock_summary.asp?menupos="& menupos &"&dtype=mk&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype &"&stplace="&stplace &"&shopid="&ojaego.FItemList(i).Fshopid&"&etcjungsantype="&etcjungsantype & "&showDiff="&showDiff
if Not(isOnlySys) THEN iURL=iURL&"&sysorreal="& sysorreal

iURLEtc = "monthlystock_etcChulgoList.asp?menupos="& menupos &"&dtype=mk&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
iURLEtc = iURLEtc + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype &"&stplace="&stplace &"&shopid="&shopid&"&etcjungsantype="&etcjungsantype
if Not(isOnlySys) THEN iURLEtc=iURLEtc&"&sysorreal="& sysorreal
%>
<% if sysorreal="real" then %>
	<tr align="right" bgcolor="<%=CHKIIF((isItemList or isGroupByBrand) and ojaego.FItemList(i).getCalcuCurRealStock<>ojaego.FItemList(i).FTotRealStockCount,"yellow","#FFFFFF")%>" >
<% else %>
	<tr align="right" bgcolor="<%=CHKIIF((isItemList or isGroupByBrand) and ojaego.FItemList(i).getCalcuCurSysStock<>ojaego.FItemList(i).FTotCount,"yellow","#FFFFFF")%>" >
<% end if %>
    <% if (isGroupByBrand) then %>
	    <td align="center">
			<a href="javascript:fnResearchByBrand('<% if (ojaego.FItemList(i).Fmakerid <> "") then %><%= ojaego.FItemList(i).Fmakerid %><% else %>-<% end if %>');">
				<% if (ojaego.FItemList(i).Fmakerid <> "") then %><%= ojaego.FItemList(i).Fmakerid %><% else %>-<% end if %>
			</a>
		</td>
		<td align="center"><%= ojaego.FItemList(i).fpurchasetypename %></td>
    <% else %>
	    <td align="center">
	        <% if makerid<>"" then%>
				<%= ojaego.FItemList(i).Fshopid %>
	        <% else %>
				<%= ojaego.FItemList(i).getBusiName %>
	        <% end if %>
	    </td>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= GetItemGubunName(ojaego.FItemList(i).Fitemgubun) %></a></td>
		<td align="center">
		    <% if makerid<>"" then%>
				<a href="javascript:TnPopItemStockWithGubun('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>','<%= ojaego.FItemList(i).Fshopid %>')">
				<%= ojaego.FItemList(i).getLogisticsCode%></a>
		    <% else %>
				<a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).Fitemgubun %></a>
		    <% end if %>
		</td>
		<td align="left">
		    <%= ojaego.FItemList(i).fitemname %>
		</td>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
		<td align="center">
			<% if (makerid = "") then %>
			<%= ojaego.FItemList(i).Fshopid %>
			<% else %>
			<%= ojaego.FItemList(i).FlastIpgoDate %>
			<% end if %>
		</td>
	<% end if %>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber(ojaego.FItemList(i).FTotPreRealStockCount,0) %>
		<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FTotPreCount,0) %>
		<% end if %>
	</td>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber(ojaego.FItemList(i).FTotPreRealStockBuySum,0) %>
		<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FTotPreBuySum,0) %>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotIpCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotIpBuySum,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotMoveItemCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotMoveItemBuySum,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotSellCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotSellBuySum,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotOffChulCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotOffChulBuySum,0) %></td>

	<% if makerid<>"" then%>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotEtcChulCount,0) %></td>
	    <td><%= FormatNumber(ojaego.FItemList(i).FTotEtcChulBuySum,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></td>
	<% else %>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotEtcChulCount,0) %></td>
	    <td><%= FormatNumber(ojaego.FItemList(i).FTotEtcChulBuySum,0) %></td>
		<td><a href="<%= iURLEtc %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></a></td>
		<td><a href="<%= iURLEtc %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></a></td>
	<% end if %>

	<td><%= FormatNumber(ojaego.FItemList(i).FTotCsChulCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotCsChulBuySum,0) %></td>
	<%
	if sysorreal="real" then
		diffStock=diffStock+(ojaego.FItemList(i).getCalcuCurRealStock-ojaego.FItemList(i).FTotRealStockCount)*-1
		diffStockPrc=diffStockPrc+(ojaego.FItemList(i).getCalcuCurRealBuySum-ojaego.FItemList(i).FTotRealStockBuySum)*-1
	else
		diffStock=diffStock+(ojaego.FItemList(i).getCalcuCurSysStock-ojaego.FItemList(i).FTotCount)*-1
		diffStockPrc=diffStockPrc+(ojaego.FItemList(i).getCalcuCurSysBuySum-ojaego.FItemList(i).FTotBuySum)*-1
	end if
	%>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber((ojaego.FItemList(i).getCalcuCurRealStock-ojaego.FItemList(i).FTotRealStockCount)*-1,0) %>
		<% else %>
			<%= FormatNumber((ojaego.FItemList(i).getCalcuCurSysStock-ojaego.FItemList(i).FTotCount)*-1,0) %>
		<% end if %>
	</td>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber((ojaego.FItemList(i).getCalcuCurRealBuySum-ojaego.FItemList(i).FTotRealStockBuySum)*-1,0) %>
		<% else %>
			<%= FormatNumber((ojaego.FItemList(i).getCalcuCurSysBuySum-ojaego.FItemList(i).FTotBuySum)*-1,0) %>
		<% end if %>
	</td>
    <td>
		<% if sysorreal="real" then %>
			<%= FormatNumber(ojaego.FItemList(i).FTotRealStockCount,0) %>
		<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %>
		<% end if %>
    </td>
	<td>
		<% if sysorreal="real" then %>
			<%= FormatNumber(ojaego.FItemList(i).FTotRealStockBuySum,0) %>
		<% else %>
			<%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %>
		<% end if %>
	</td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotErrRealCheckCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotErrRealCheckBuySum,0) %></td>
	<% if sysorreal="sys" then %>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotRealStockCount,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotRealStockBuySum,0) %></td>
	<% end if %>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotErrBadItemCount,0) %></td>
	<td><%= FormatNumber(ojaego.FItemList(i).FTotErrBadItemBuySum,0) %></td>
	<% if (isViewWonga) then %>
		<td>
			<%= FormatNumber(ojaego.FItemList(i).getWongaCnt,0) %>
			<% if sysorreal="real" then %>
				<% if (ojaego.FItemList(i).getCalcuCurRealStock<>ojaego.FItemList(i).FTotPreRealStockCount) and (showDiff = "Y") then %>
					<% diffStockW=diffStockW+ojaego.FItemList(i).FTotPreRealStockCount-ojaego.FItemList(i).getCalcuCurRealStock %>
					<br><font color="red"><%= FormatNumber(ojaego.FItemList(i).FTotPreRealStockCount-ojaego.FItemList(i).getCalcuCurRealStock,0) %></font>
				<% end if %>
			<% else %>
				<% if (ojaego.FItemList(i).getCalcuCurSysStock<>ojaego.FItemList(i).FTotCount) and (showDiff = "Y") then %>
					<% diffStockW=diffStockW+ojaego.FItemList(i).FTotCount-ojaego.FItemList(i).getCalcuCurSysStock %>
					<br><font color="red"><%= FormatNumber(ojaego.FItemList(i).FTotCount-ojaego.FItemList(i).getCalcuCurSysStock,0) %></font>
				<% end if %>
			<% end if %>
		</td>
		<td><%= FormatNumber(ojaego.FItemList(i).getWongaSum,0) %></td>
		<td align="center" valign="middle">
			<% if makerid<>"" then %>
				<% if sysorreal="real" then %>
					<% if (isItemList or isGroupByBrand) and ojaego.FItemList(i).getCalcuCurRealStock<>ojaego.FItemList(i).FTotPreRealStockCount then %>
						<a href="javascript:reActAccMonthSummaryOneItem('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>')"><img src="/images/icon_reload.gif" border="0"></a>
					<% end if %>
				<% else %>
					<% if (isItemList or isGroupByBrand) and ojaego.FItemList(i).getCalcuCurSysStock<>ojaego.FItemList(i).FTotCount then %>
						<a href="javascript:reActAccMonthSummaryOneItem('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>')"><img src="/images/icon_reload.gif" border="0"></a>
					<% end if %>
				<% end if %>
			<% end if %>
		</td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaCnt,0) %></td>
		<td><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaSum,0) %></td>
	<% end if %>
	<% if (isItemList) then %>
		<td align="center"><img src="/images/icon_arrow_link.gif" style="cursor:pointer" onClick="popAccStockModiOne('<%= ojaego.FItemList(i).FItemgubun %>', '<%= ojaego.FItemList(i).FItemid %>', '<%= ojaego.FItemList(i).FItemOption %>')"></td>
    <% end if %>
</tr>
<%
end if
next
%>
<% if (Not isGroupByBrand) then %>
	<tr align="right" bgcolor="#EEFFEE">
		<td></td>
		<td align="center">저장품소계</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td>
			<% if sysorreal="real" then %>
				<%= FormatNumber(subPreRealStockno,0) %>
			<% else %>
				<%= FormatNumber(subPreno,0) %>
			<% end if %>
		</td>
		<td>
			<% if sysorreal="real" then %>
				<%= FormatNumber(subPreRealStockbuy,0) %>
			<% else %>
				<%= FormatNumber(subPrebuy,0) %>
			<% end if %>
		</td>
		<td><%= FormatNumber(subIpno,0) %></td>
		<td><%= FormatNumber(subIpBuy,0) %></td>
		<td><%= FormatNumber(subMoveItemno,0) %></td>
		<td><%= FormatNumber(subMoveItemBuy,0) %></td>
		<td><%= FormatNumber(subSellno,0) %></td>
		<td><%= FormatNumber(subSellBuy,0) %></td>
		<td><%= FormatNumber(subOffChulno,0) %></td>
		<td><%= FormatNumber(subOffChulBuy,0) %></td>
		<td><%= FormatNumber(subEtcChulno,0) %></td>
		<td><%= FormatNumber(subEtcChulBuy,0) %></td>
		<td><%= FormatNumber(subLossno,0) %></td>
		<td><%= FormatNumber(subLossBuy,0) %></td>
		<td><%= FormatNumber(subCsChulno,0) %></td>
		<td><%= FormatNumber(subCsChulBuy,0) %></td>
		<td><%= FormatNumber(diffStock,0) %></td>
		<td><%= FormatNumber(diffStockPrc,0) %></td>
		<td>
			<% if sysorreal="real" then %>
				<%= FormatNumber(subRealStockno,0) %>
			<% else %>
				<%= FormatNumber(subTotno,0) %>
			<% end if %>
		</td>
		<td>
			<% if sysorreal="real" then %>
				<%= FormatNumber(subRealStockBuy,0) %>
			<% else %>
				<%= FormatNumber(subTotbuy,0) %>
			<% end if %>
		</td>
		<td><%= FormatNumber(subErrRealCheckno,0) %></td>
		<td><%= FormatNumber(subErrRealCheckBuy,0) %></td>
		<% if sysorreal="sys" then %>
			<td><%= FormatNumber(subRealStockno,0) %></td>
			<td><%= FormatNumber(subRealStockBuy,0) %></td>
		<% end if %>
		<td><%= FormatNumber(subErrBadItemno,0) %></td>
		<td><%= FormatNumber(subErrBadItemBuy,0) %></td>
		<% if (isViewWonga) then %>
			<td>
				<% if sysorreal="real" then %>
					<%= FormatNumber(subPreRealStockno+subIpno-subRealStockno,0) %>
				<% else %>
					<%= FormatNumber(subPreno+subIpno-subtotno,0) %>
				<% end if %>
			</td>
			<td>
				<% if sysorreal="real" then %>
					<%= FormatNumber(subPreRealStockbuy+subIpBuy-subRealStockBuy,0) %>
				<% else %>
					<%= FormatNumber(subPrebuy+subIpBuy-subtotbuy,0) %>
				<% end if %>
			</td>
			<td></td>
			<td><%= FormatNumber(subLossno,0) %></td>
			<td><%= FormatNumber(subLossBuy,0) %></td>
			<td>
				<% if sysorreal="real" then %>
					<%= FormatNumber(subPreRealStockno+subIpno-subRealStockno+subLossno,0) %>
				<% else %>
					<%= FormatNumber(subPreno+subIpno-subtotno+subLossno,0) %>
				<% end if %>
			</td>
			<td>
				<% if sysorreal="real" then %>
					<%= FormatNumber(subPreRealStockbuy+subIpBuy-subRealStockBuy+subLossBuy,0) %>
				<% else %>
					<%= FormatNumber(subPrebuy+subIpBuy-subtotbuy+subLossBuy,0) %>
				<% end if %>
			</td>
		<% end if %>
		<% if (isItemList) then %>
			<td></td>
		<% end if %>
	</tr>
	<!--<tr  bgcolor="#FFFFFF"><td colspan="32"></td></tr>-->
	<%
	subTotno=0
	subTotbuy=0
	subPreno   =0
	subPrebuy  =0
	subPreRealStockno   =0
	subPreRealStockbuy  =0
	subIpno    =0
	subIpBuy    =0
	subLossno   =0
	subLossBuy  =0

	subSellno       = 0
	subSellBuy      = 0
	subOffChulno    = 0
	subOffChulBuy   = 0
	subEtcChulno    = 0
	subEtcChulBuy   = 0
	subCsChulno     = 0
	subCsChulBuy    = 0
	subErrBadItemno = 0
	subErrBadItemBuy = 0
	subErrRealCheckno = 0
	subErrRealCheckBuy = 0
	subRealStockno = 0
	subRealStockBuy = 0
	subMoveItemno = 0
	subMoveItemBuy = 0
	%>
<% end if %>
<tr align="center" bgcolor="#FFFFFF">
    <% if (isGroupByBrand) then %>
		<td colspan="2">총계</td>
    <% else %>
		<td></td>
		<td>총계</td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	<% end if %>
	<td align="right" >
		<% if sysorreal="real" then %>
			<%= FormatNumber(totPreRealStockno,0) %>
		<% else %>
			<%= FormatNumber(totPreno,0) %>
		<% end if %>
	</td>
	<td align="right" >
		<% if sysorreal="real" then %>
			<%= FormatNumber(totPreRealStockbuy,0) %>
		<% else %>
			<%= FormatNumber(totPrebuy,0) %>
		<% end if %>
	</td>
	<td align="right" ><%= FormatNumber(totIpno,0) %></td>
	<td align="right" ><%= FormatNumber(totIpBuy,0) %></td>
	<td align="right" ><%= FormatNumber(totMoveItemno,0) %></td>
	<td align="right" ><%= FormatNumber(totMoveItemBuy,0) %></td>
	<td align="right" ><%= FormatNumber(totSellno,0) %></td>
	<td align="right" ><%= FormatNumber(totSellBuy,0) %></td>
	<td align="right" ><%= FormatNumber(totOffChulno,0) %></td>
	<td align="right" ><%= FormatNumber(totOffChulBuy,0) %></td>
	<td align="right" ><%= FormatNumber(totEtcChulno,0) %></td>
	<td align="right" ><%= FormatNumber(totEtcChulBuy,0) %></td>
	<td align="right" ><%= FormatNumber(totLossno,0) %></td>
	<td align="right" ><%= FormatNumber(totLossBuy,0) %></td>
	<td align="right" ><%= FormatNumber(totCsChulno,0) %></td>
	<td align="right" ><%= FormatNumber(totCsChulBuy,0) %></td>
	<td align="right" ><%= FormatNumber(diffStock,0) %></td>
	<td align="right" ><%= FormatNumber(diffStockPrc,0) %></td>
	<td align="right" >
		<% if sysorreal="real" then %>
			<%= FormatNumber(totRealStockno,0) %>
		<% else %>
			<%= FormatNumber(totno,0) %>
		<% end if %>
	</td>
	<td align="right" >
		<% if sysorreal="real" then %>
			<%= FormatNumber(totRealStockBuy,0) %>
		<% else %>
			<%= FormatNumber(totbuy,0) %>
		<% end if %>
	</td>
	<td align="right" ><%= FormatNumber(totErrRealCheckno,0) %></td>
	<td align="right" >
		<%= FormatNumber(totErrRealCheckBuy,0) %>
		<% if (showDiff = "Y") and ((isGroupByBrand) or (makerid <> "")) then %>
		<br /><font color="red">(+<%= FormatNumber(totErrRealCheckBuyPlus,0) %>)<br />(<%= FormatNumber(totErrRealCheckBuyMinus,0) %>)</font>
		<% end if %>
	</td>
	<% if sysorreal="sys" then %>
		<td align="right" ><%= FormatNumber(totRealStockno,0) %></td>
		<td align="right" ><%= FormatNumber(totRealStockBuy,0) %></td>
	<% end if %>
	<td align="right" ><%= FormatNumber(totErrBadItemno,0) %></td>
	<td align="right" ><%= FormatNumber(totErrBadItemBuy,0) %></td>
	<% if (isViewWonga) then %>
		<td align="right" >
			<% if sysorreal="real" then %>
				<%= FormatNumber(totPreRealStockno+totIpno-totRealStockno,0) %>
			<% else %>
				<%= FormatNumber(totPreno+totIpno-totno,0) %>
			<% end if %>
			
	    	<% if (diffStockW<>0) then %>
				<br><font color=red><%= FormatNumber(diffStockW,0) %></font>
	    	<% end if %>
		</td>
		<td align="right" >
			<% if sysorreal="real" then %>
				<%= FormatNumber(totPreRealStockbuy+totIpBuy-totRealStockBuy,0) %>
			<% else %>
				<%= FormatNumber(totPrebuy+totIpBuy-totbuy,0) %>
			<% end if %>
		</td>
		<td ></td>
		<td align="right" ><%= FormatNumber(totLossno,0) %></td>
		<td align="right" ><%= FormatNumber(totLossBuy,0) %></td>
		<td align="right" >
			<% if sysorreal="real" then %>
				<%= FormatNumber(totPreRealStockno+totIpno-totRealStockno+totLossno,0) %>
			<% else %>
				<%= FormatNumber(totPreno+totIpno-totno+totLossno,0) %>
			<% end if %>
		</td>
		<td align="right" >
			<% if sysorreal="real" then %>
				<%= FormatNumber(totPreRealStockbuy+totIpBuy-totRealStockBuy+totLossBuy,0) %>
			<% else %>
				<%= FormatNumber(totPrebuy+totIpBuy-totbuy+totLossBuy,0) %>
			<% end if %>
		</td>
	<% end if %>
	<% if (isItemList) then %>
		<td ></td>
    <% end if %>
</tr>
</table>

<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
