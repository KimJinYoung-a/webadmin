<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인매장 브랜드별 재고 파악
' History : 2011.08.01 이상구 생성
'			2019.05.31 한용민 수정
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopitemcls.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->

<%
dim shopid, makerid, research
	shopid       = RequestCheckVar(request("shopid"),32)
	makerid      = RequestCheckVar(request("makerid"),32)
	research     = RequestCheckVar(request("research"),32)

dim usingyn, centermwdiv ,NoZeroStock, comm_cd
	usingyn      = RequestCheckVar(request("usingyn"),32)
	centermwdiv  = RequestCheckVar(request("centermwdiv"),32)
	NoZeroStock  = RequestCheckVar(request("NoZeroStock"),32)
	comm_cd      = RequestCheckVar(request("comm_cd"),32)

'/매장
if (C_IS_SHOP) then

	'//직영점일때
	if C_IS_OWN_SHOP then

		'/어드민권한 점장 미만
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/업체
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")	'"7321"
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

if (research="") then NoZeroStock="on"

dim oOffStock
set oOffStock = new CShopItemSummary
	oOffStock.FRectShopID       = shopid
	oOffStock.FRectMakerID      = makerid
	oOffStock.FRectComm_cd      = comm_cd
	oOffStock.FRectIsUsing      = usingyn
	oOffStock.FRectNoZeroStock  = NoZeroStock

	if (shopid<>"") then
	    oOffStock.GetShopBrandRealCheckRequire
	end if

dim i
dim sumTotItemNo            : sumTotItemNo=0
dim sumStPLusStockItemCnt   : sumStPLusStockItemCnt=0
dim sumTotSellNo            : sumTotSellNo=0
dim sumTotRealStockNo       : sumTotRealStockNo=0
dim sumTotStockBuySum       : sumTotStockBuySum=0
dim sumTotOwnStockBuySum    : sumTotOwnStockBuySum=0

%>

<script language='javascript'>

function popBrandStock(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&research=on"+"&NoZeroStock=on";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandStockTaking(shopid,makerid){
    var popUrl = "/common/offshop/shop_brandcurrentstock_takingWithList.asp?menupos=1074&shopid="+shopid+"&makerid="+makerid+"&research=on"+"&NoZeroStock=on";
    var popwin = window.open(popUrl,'popBrandStock','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandStockTakingInput(stIdx){
    var popUrl = "/common/offshop/shop_brandcurrentstock_byjobkey.asp?idx="+stIdx+"&sType=stTaking";
    var popwin = window.open(popUrl,'popBrandStockInput','scrollbars=yes,resizable=yes');
    popwin.focus();
}

function frmsumbit(page){
	frm.page.value=page;
	frm.action="";
	frm.target = "";
	frm.submit();
}

function jsCurrStockDown(stockPlace,temp){
	if (stockPlace==""){
		alert('재고위치가 지정되지 않았습니다.');
		return;
	}
	frm.stockPlace.value=stockPlace;
	frm.action="/admin/newreport/currentstock_excel.asp";
	frm.target = "view";
	frm.submit();
	frm.target = "";
	frm.action = ""
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<input type="hidden" name="page" value="">
<input type="hidden" name="stockPlace" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		<%
		'직영/가맹점
		if (C_IS_SHOP) then
		%>
			<% if getoffshopdiv(shopid) <> "1" and shopid <> "" then %>
				<input type="hidden" name="shopid" value="<%= shopid %>">
				* 매장 : <%= shopid %>
				&nbsp;
				* 브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
			<% else %>
	        	* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
				&nbsp;
				* 브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		<%
			end if
		else
			''업체인경우
			if (C_IS_Maker_Upche) then
		%>
				* 매장 : <% drawSelectBoxOpenOffShop "shopid",shopid %>
				<input type="hidden" name="makerid" value="<%= makerid %>">
		<%
			else
				if (C_ADMIN_USER) then
		%>
					* 매장 : <% drawSelectBoxOffShop "shopid",shopid %>
					&nbsp;
					* 브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>
		<%
				end if
			end if
		end if
		%>
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="frmsumbit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		* 상품 사용구분 : <% drawSelectBoxUsingYN "usingyn", usingyn %>
		&nbsp;
	    * 매장매입구분 : <% drawSelectBoxOFFJungsanCommCDmulti "comm_cd",comm_cd %>
		&nbsp;
		<input type="checkbox" name="NoZeroStock" <%= CHKIIF(NoZeroStock="on","checked","") %> > 재고0인 브랜드 검색 안함.
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" >
<tr>
	<td align="left">
		* 마이너스 재고는 재고금액 0 으로 산정함.
	</td>
	<td align="right">
		* 재고실사(검색조건에 매장을 선택하면 해당매장 재고가 다운로드됩니다.) :
		<!--
		<br><br><input type="checkbox" name="day1after">말일이후변동값포함
		<input type="button" class="button" value="재고실사다운로드(매장)" onclick="jsstockDown('S','');">
		-->
		<input type="button" class="button" value="현재재고다운로드(<%= CHKIIF(shopid="", "streetshop011", shopid) %>)" onclick="jsCurrStockDown('S','');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="14">
		검색결과 총 <%= oOffStock.FTotalCount %> 건
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td width="20"></td>
    <td width="110">브랜드ID</td>
	<td width="100">매입구분</td>
	<td width="70">상품품목수</td>
	<td width="70">재고>0상품수</td>
	<td width="70">총판매</td>
	<td width="70">총 실사재고</td>
	<td width="90">재고액
    	<% IF (C_IS_FRN_SHOP) then %>
    		<br>(매장매입가)
    	<% else %>
    		<br>(본사매입가)
    	<% end if %>
	</td>
	<td width="90">최초 입고일</td>
	<td width="90">최종 입고일</td>
    <td >최근실사일</td>
    <td >시트(수기)<br>재고파악</td>
    <td >바코드<br>재고파악</td>
    <td >재고입력</td>
</tr>
<% if (shopid="") then %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td colspan="14">[먼저 매장 을 선택하세요.]</td>
</tr>
<% else %>
<% for i=0 to oOffStock.FResultCount-1 %>
<%
sumTotItemNo            =sumTotItemNo+oOffStock.FItemList(i).FtotItemNo
sumStPLusStockItemCnt   =sumStPLusStockItemCnt+oOffStock.FItemList(i).FstPLusStockItemCnt
sumTotSellNo            =sumTotSellNo+oOffStock.FItemList(i).FtotSellNo*-1
sumTotRealStockNo       =sumTotRealStockNo+oOffStock.FItemList(i).FtotRealStockNo
if not isNULL(oOffStock.FItemList(i).FtotStockBuySum) then

    IF (C_IS_FRN_SHOP) then
        sumTotStockBuySum       =sumTotStockBuySum+oOffStock.FItemList(i).FtotStockBuySum
    else
        sumTotOwnStockBuySum    =sumTotOwnStockBuySum+oOffStock.FItemList(i).FtotOwnStockBuySum
    end if
end if
%>
<tr align="center" bgcolor="#FFFFFF">
    <td></td>
    <td><%= oOffStock.FItemList(i).Fmakerid %></td>
    <td><%= oOffStock.FItemList(i).Fcomm_name %></td>
    <td><%= FormatNumber(oOffStock.FItemList(i).FtotItemNo,0) %></td>
    <td><%= FormatNumber(oOffStock.FItemList(i).FstPLusStockItemCnt,0) %></td>
    <td><%= FormatNumber(oOffStock.FItemList(i).FtotSellNo*-1,0) %></td>
    <td><%= FormatNumber(oOffStock.FItemList(i).FtotRealStockNo,0) %></td>
    <td align="right">
	    <% if isNULL(oOffStock.FItemList(i).FtotStockBuySum) then %>
			<% if (UBound(Split(oOffStock.FItemList(i).Fmakerid, "-")) = 2) then %>
				<font color=red>상품정보 없음</font>
			<% else %>
				<font color=red>계약 없음</font>
			<% end if %>
	    <% else %>
	        <% IF (C_IS_FRN_SHOP) then %>
	        <%= FormatNumber(oOffStock.FItemList(i).FtotStockBuySum,0) %>
	        <% else %>
	        <%= FormatNumber(oOffStock.FItemList(i).FtotOwnStockBuySum,0) %>
	        <% end if %>
	    <% end if %>
    </td>
    <td><%= oOffStock.FItemList(i).Ffirstipgodate %></td>
    <td><%= oOffStock.FItemList(i).Flastipgodate %></td>
    <td><%= oOffStock.FItemList(i).FlastStockdate %></td>
    <td><input type="button" class="button" value="수기 입력" onClick="popBrandStock('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');"></td>
    <td>
	    <% if oOffStock.FItemList(i).FstStatus=0 then %>
			<input type="button" class="button_ing" value="재고 파악 中" onClick="popBrandStockTaking('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');">
	    <% elseif oOffStock.FItemList(i).FstStatus=3 then %>
			<input type="button" class="button" disabled value="재고 파악" onClick="popBrandStockTaking('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');">
	    <% else %>
			<input type="button" class="button" value="재고 파악" onClick="popBrandStockTaking('<%= shopid %>','<%= oOffStock.FItemList(i).Fmakerid %>');">
	    <% end if %>
    </td>
    <td>
	    <% if oOffStock.FItemList(i).FstStatus=3 then %>
			<input type="button" class="button_ing" value="재고 입력" onClick="popBrandStockTakingInput(<%= oOffStock.FItemList(i).FstTakingIdx %>);">
	    <% else %>
			<input type="button" class="button" disabled value="재고 입력" onClick="popBrandStockTakingInput(<%= oOffStock.FItemList(i).FstTakingIdx %>);">
	    <% end if %>
    </td>
</tr>
<% next %>
<tr align="center" bgcolor="#FFFFFF">
    <td></td>
    <td></td>
    <td></td>
    <td><%= formatNumber(sumTotItemNo,0) %></td>
    <td><%= formatNumber(sumStPLusStockItemCnt,0) %></td>
    <td><%= formatNumber(sumTotSellNo,0) %></td>
    <td><%= formatNumber(sumTotRealStockNo,0) %></td>

    <% IF (C_IS_FRN_SHOP) then %>
		<td align="right"><%= formatNumber(sumTotStockBuySum,0) %></td>
    <% else %>
		<td align="right"><%= formatNumber(sumTotOwnStockBuySum,0) %></td>
    <% end if %>

    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
</tr>
<% end if %>
</table>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe id="view" name="view" src="" width="100%" height="300" frameborder="0" scrolling="no"></iframe>
<% else %>
	<iframe id="view" name="view" src="" width="100%" height="0" frameborder="0" scrolling="no"></iframe>
<% end if %>
<%
set oOffStock = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
