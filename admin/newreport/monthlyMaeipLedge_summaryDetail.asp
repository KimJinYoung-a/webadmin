<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  재고
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls.asp"-->

<%
Dim CCADMIN : CCADMIN = C_ADMIN_AUTH
dim showItemDetailPopup : showItemDetailPopup = (left(now, 10) <= "2014-08-09") or CCADMIN


dim research, i, page
dim yyyy1,mm1, yyyymm1, makerid, showsuply, meaipTp, showShopid, showDiff
dim stockPlace, shopid
dim targetGbn, itemgubun
dim stype       '' S:재고, J:정산
dim bPriceGbn
dim showUpbae
dim PurchaseType
dim showPoint

page        = requestCheckvar(request("page"),10)
stype       = requestCheckvar(request("stype"),10)
shopid    	= requestCheckvar(request("shopid"),32)
research    = requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1       	= requestCheckvar(request("mm1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsuply   = requestCheckvar(request("showsuply"),10)
showShopid  = requestCheckvar(request("showShopid"),10)
meaipTp     = requestCheckvar(request("meaipTp"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
targetGbn   = requestCheckvar(request("targetGbn"),10)
showDiff   	= requestCheckvar(request("showDiff"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
showUpbae   = requestCheckvar(request("showUpbae"),10)
PurchaseType   = requestCheckvar(request("PurchaseType"),10)
showPoint   = requestCheckvar(request("showPoint"),10)

if (page="") then page=1
if (stockPlace="") then stockPlace="L"

if yyyy1 = "" then
	yyyy1 = "2013"
	mm1 = "12"
end if

if (research="") and (bPriceGbn = "") then
    bPriceGbn="P"
end if

yyyymm1 = yyyy1 + "-" + mm1

dim oCMonthlyMaeipLedge
set oCMonthlyMaeipLedge = new CMonthlyMaeipLedge

oCMonthlyMaeipLedge.FRectYYYYMM = yyyymm1
oCMonthlyMaeipLedge.FRectStockPlace = stockPlace
oCMonthlyMaeipLedge.FRectShopid = shopid
oCMonthlyMaeipLedge.FRectMakerid = makerid
oCMonthlyMaeipLedge.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
oCMonthlyMaeipLedge.FRectMeaipTp = meaipTp
oCMonthlyMaeipLedge.FRectItemgubun = itemgubun
oCMonthlyMaeipLedge.FRectTargetGbn = targetGbn
oCMonthlyMaeipLedge.FRectShowShopid = showShopid
oCMonthlyMaeipLedge.FRectPriceGubun = bPriceGbn
oCMonthlyMaeipLedge.FRectShowUpbae = showUpbae

oCMonthlyMaeipLedge.FRectShowDiff = showDiff
oCMonthlyMaeipLedge.FRectShowPurchaseType = "Y"
oCMonthlyMaeipLedge.FRectPurchaseType = PurchaseType
oCMonthlyMaeipLedge.FRectShowPoint = showPoint

oCMonthlyMaeipLedge.FPageSize = 4000
oCMonthlyMaeipLedge.FCurrPage = page

if (stype="S") then
    oCMonthlyMaeipLedge.GetMaeipLedgeSUMSubDetail
else
    oCMonthlyMaeipLedge.GetMaeipJungsanSumSubDetail
end if


dim totprevSysStockNo, totprevSysStockSum, totIpgoNo, totIpgoSum, totSellNo, totSellSum, totOffChulNo, totOffChulSum, totEtcChulNo, totEtcChulSum
dim totCsNo, totCsSum, totLossChulNo, totLossChulSum, totcurSysStockNo, totcurSysStockSum, totcurErrRealCheckNo, totcurErrRealCheckSum
dim diff, totdiff
dim diffSum, totdiffSum
dim totMoveNo, totMoveSum
dim totErrNo, totErrSum

%>
<script language='javascript'>

function NextPage(page){
	document.frm.page.value = page;
	document.frm.submit();
}

function GotoBrand(makerid){
	document.frm.makerid.value = makerid;
	document.frm.submit();
}

<% if (showItemDetailPopup = True) then %>
function PopItemStock(itemgubun, itemid, itemoption) {
	var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"PopItemStock","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
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

	var popwin = window.open("/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=" + shopid + "&barcode=" + barcode,"PopItemStockShop","width=1000 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
<% end if %>

function popAccStockModiOne(itemgubun,itemid,itemoption){
	var popwin = window.open("/admin/newreport/pop_item_stock_Accsummary_edit.asp?yyyy1=2015&mm1=03&shopid=&itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"popAccStockModiOne","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="" target="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="<%= page %>">
	<input type="hidden" name="stype" value="<%=stype%>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %>
			&nbsp;&nbsp;
			<font color="#CC3333">브랜드:</font> <%	drawSelectBoxDesignerWithName "makerid", makerid %>
			&nbsp;&nbsp;
			매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showsuply" value="on" <%= CHKIIF(showsuply="on","checked","") %> >공급가로 표시
			&nbsp;&nbsp;
			<input type="checkbox" name="showShopid" value="on" <%= CHKIIF(showShopid="on","checked","") %> >매장 표시
			<% if (CCADMIN) then %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showDiff" value="on" <%= CHKIIF(showDiff="on","checked","") %> > 오차내역만 표시
		    <% end if %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showUpbae" value="on" <%= CHKIIF(showUpbae="on","checked","") %> >업배상품만 표시
            &nbsp;&nbsp;
			<input type="checkbox" name="showPoint" value="on" <%= CHKIIF(showPoint="on","checked","") %> >소수점상품 표시
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.target='';document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <font color="#CC3333">재고위치:</font>
		    <select name="stockPlace">
        		<option value="L" <%= CHKIIF(stockPlace="L","selected" ,"") %> >물류</option>
        		<option value="S" <%= CHKIIF(stockPlace="S","selected" ,"") %> >매장</option>
				<option value="T" <%= CHKIIF(stockPlace="T","selected" ,"") %> >띵소</option>
				<option value="O" <%= CHKIIF(stockPlace="O","selected" ,"") %> >온라인정산</option>
				<option value="N" <%= CHKIIF(stockPlace="N","selected" ,"") %> >온라인 매입정산(공제불가)</option>
				<option value="F" <%= CHKIIF(stockPlace="F","selected" ,"") %> >오프정산</option>
				<option value="A" <%= CHKIIF(stockPlace="A","selected" ,"") %> >핑거스정산</option>
				<option value="E" <%= CHKIIF(stockPlace="E","selected" ,"") %> >에러</option>
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">매입구분:</font>
        	<select name="meaipTp">
        	<option value="">전체
        	<option value="M" <%= CHKIIF(meaipTp="M","selected" ,"") %> >입고분매입
        	<option value="S" <%= CHKIIF(meaipTp="S","selected" ,"") %> >판매분매입
        	<option value="C" <%= CHKIIF(meaipTp="C","selected" ,"") %> >출고분매입
        	<option value="E" <%= CHKIIF(meaipTp="E","selected" ,"") %> >기타매입
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">부서구분:</font>
        	<input type="text" name="targetGbn" value="<%=targetGbn%>" size="2" maxlength="2">

        	&nbsp;&nbsp;
        	<font color="#CC3333">코드구분:</font>
        	<input type="text" name="itemgubun" value="<%=itemgubun%>" size="2" maxlength="2">
        	&nbsp;&nbsp;
			<font color="#CC3333">매입가기준:</font>
			<input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >작성시매입가
			<input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %>  >평균매입가
			&nbsp;&nbsp;
			<font color="#CC3333">구매유형:</font>
			<% drawPartnerCommCodeBox true,"purchasetype","PurchaseType",PurchaseType,"" %>
	    </td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

* 현재달에 입출고 및 재고가 없더라도, 이전달에 입출고 또는 재고가 있는 경우 표시됩니다.

<p />

<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF">
<tr>
<td>총 <%=oCMonthlyMaeipLedge.FTotalCount%> 건 <%=page%>/<%=oCMonthlyMaeipLedge.FtotalPage%> page</td>
</tr>
</table>
<p>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="6">상품구분</td>
        <td colspan="2">기초재고(월말일자)</td>
        <td colspan="2">당월매입(월)</td>
        <td colspan="2">당월이동(월)</td>
        <td colspan="2">당월판매(월)</td>
        <td colspan="2">당월출고1(월)</td>
        <td colspan="2">당월출고2(월)</td>
        <td colspan="2">당월기타출고(월)</td>
        <td colspan="2">당월CS출고(월)</td>
        <td colspan="2">오차(월)</td>
		<td colspan="2"><b>기말재고(월)</b></td>
		<td rowspan="2">비고검토</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td >매장ID</td>
		<td >브랜드ID</td>
	    <td >상품코드</td>
		<td >구매유형</td>
	    <td >매입구분</td>
	    <td >재고위치</td>
    	<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
		<td>수량</td>
		<td>금액</td>
    </tr>
    <% for i=0 to oCMonthlyMaeipLedge.FResultCount-1 %>
    <%

	totprevSysStockNo       	= totprevSysStockNo + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo
	totprevSysStockSum       	= totprevSysStockSum + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum

	totIpgoNo       			= totIpgoNo + oCMonthlyMaeipLedge.FItemList(i).getIpgoNo
	totIpgoSum       			= totIpgoSum + Round(oCMonthlyMaeipLedge.FItemList(i).getIpgoSum,0)

	totMoveNo       			= totMoveNo + oCMonthlyMaeipLedge.FItemList(i).getMoveNo
	totMoveSum       			= totMoveSum + oCMonthlyMaeipLedge.FItemList(i).getMoveSum

	totSellNo       			= totSellNo + oCMonthlyMaeipLedge.FItemList(i).FSellNo
	totSellSum       			= totSellSum + oCMonthlyMaeipLedge.FItemList(i).FSellSum

	totOffChulNo       			= totOffChulNo + oCMonthlyMaeipLedge.FItemList(i).FOffChulNo
	totOffChulSum       		= totOffChulSum + oCMonthlyMaeipLedge.FItemList(i).FOffChulSum

	totEtcChulNo       			= totEtcChulNo + oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo
	totEtcChulSum       		= totEtcChulSum + oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum

	totLossChulNo       		= totLossChulNo + oCMonthlyMaeipLedge.FItemList(i).FLossChulNo
	totLossChulSum       		= totLossChulSum + oCMonthlyMaeipLedge.FItemList(i).FLossChulSum

	totCsNo       				= totCsNo + oCMonthlyMaeipLedge.FItemList(i).FCsNo
	totCsSum       				= totCsSum + oCMonthlyMaeipLedge.FItemList(i).FCsSum

	totcurSysStockNo       		= totcurSysStockNo + oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo
	totcurSysStockSum       	= totcurSysStockSum + oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum
	totcurErrRealCheckNo       	= totcurErrRealCheckNo + oCMonthlyMaeipLedge.FItemList(i).FcurErrRealCheckNo
	totcurErrRealCheckSum       = totcurErrRealCheckSum + oCMonthlyMaeipLedge.FItemList(i).FcurErrRealCheckSum

	'diff = oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo + oCMonthlyMaeipLedge.FItemList(i).getIpgoNo + oCMonthlyMaeipLedge.FItemList(i).getMoveNo + oCMonthlyMaeipLedge.FItemList(i).FSellNo + oCMonthlyMaeipLedge.FItemList(i).FOffChulNo + oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo + oCMonthlyMaeipLedge.FItemList(i).FCsNo + oCMonthlyMaeipLedge.FItemList(i).FLossChulNo - oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo
	diff = oCMonthlyMaeipLedge.FItemList(i).getDiffNo
	diffSum = oCMonthlyMaeipLedge.FItemList(i).getDiffSum
	totdiff = totdiff + diff
	totdiffSum = totdiffSum + diffSum

	totErrNo = totErrNo + oCMonthlyMaeipLedge.FItemList(i).getTotErrNo
    totErrSum = totErrSum + oCMonthlyMaeipLedge.FItemList(i).getTotErrSum
    %>
    <tr align="right" bgcolor="#FFFFFF" onmouseover="this.style.background='F1F1F1'" onmouseout="this.style.background='FFFFFF'" >
		<td align="center">
		    <% if (showShopid<>"") then %>
		    	<%= oCMonthlyMaeipLedge.FItemList(i).Fshopid %>
			<% elseif (shopid <> "") then %>
				<%= shopid %>
		    <% end if %>
		</td>
		<td align="center"><a href="javascript:GotoBrand('<%= oCMonthlyMaeipLedge.FItemList(i).FMakerid%>')"><%= oCMonthlyMaeipLedge.FItemList(i).FMakerid%></a></td>
		<td align="center">
		    <% if (makerid<>"") then %>
				<% if (showItemDetailPopup = True) then %>
					<% if (oCMonthlyMaeipLedge.FItemList(i).Fshopid <> "") then %>
						<a href="javascript:PopItemStockShop('<%= oCMonthlyMaeipLedge.FItemList(i).Fshopid %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemid %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemoption %>');">
					<% else %>
						<a href="javascript:PopItemStock('<%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemid %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemoption %>');">
					<% end if %>
				<% end if %>
		        <%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %>-<%= oCMonthlyMaeipLedge.FItemList(i).Fitemid %>-<%= oCMonthlyMaeipLedge.FItemList(i).Fitemoption %>
		    <% else %>

		    <% end if %>
		</td>
		<td align="center"><%= getBrandPurchaseType(oCMonthlyMaeipLedge.FItemList(i).FpurchaseType) %></td>
        <td align="center"><%= oCMonthlyMaeipLedge.FItemList(i).getMeaipTypeName %></td>
        <td align="center"><%= oCMonthlyMaeipLedge.FItemList(i).FstockPlace%></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getIpgoNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getIpgoSum,0) %></td>

        <td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getMoveNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getMoveSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FSellNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FSellSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FOffChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FOffChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FCsNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FCsSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getTotErrNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getTotErrSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum,0) %></td>

		<td align="center"><img src="/images/icon_arrow_link.gif" style="cursor:pointer" onClick="popAccStockModiOne('<%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemid %>', '<%= oCMonthlyMaeipLedge.FItemList(i).Fitemoption %>')"></td>


    </tr>
    <% if FIx(i / 1000)=(i / 1000) then response.flush %>
	<% next %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
		<td></td>
		<td></td>
		<td></td>
    	<td></td>
        <td></td>
    	<td align="right" ><%= FormatNumber(totprevSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(totprevSysStockSum,0) %></td>

		<td align="right" ><%= FormatNumber(totIpgoNo,0) %></td>
		<td align="right" ><%= FormatNumber(totIpgoSum,0) %></td>

		<td align="right" ><%= FormatNumber(totMoveNo,0) %></td>
		<td align="right" ><%= FormatNumber(totMoveSum,0) %></td>

		<td align="right" ><%= FormatNumber(totSellNo,0) %></td>
		<td align="right" ><%= FormatNumber(totSellSum,0) %></td>

		<td align="right" ><%= FormatNumber(totOffChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totOffChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totEtcChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totEtcChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totLossChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totLossChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totCsNo,0) %></td>
		<td align="right" ><%= FormatNumber(totCsSum,0) %></td>

		<td align="right" ><%= FormatNumber(totErrNo,0) %></td>
		<td align="right" ><%= FormatNumber(totErrSum,0) %></td>

		<td align="right" ><%= FormatNumber(totcurSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(totcurSysStockSum,0) %></td>

		<td></td>

    </tr>

	<tr height="25" bgcolor="FFFFFF">
	    <td><%=i%></td>
		<td colspan="26" align="center">
			<% if oCMonthlyMaeipLedge.HasPreScroll then %>
        		<a href="javascript:NextPage('<%= oCMonthlyMaeipLedge.StarScrollPage-1 %>')">[pre]</a>
        	<% else %>
        		[pre]
        	<% end if %>

        	<% for i=0 + oCMonthlyMaeipLedge.StarScrollPage to oCMonthlyMaeipLedge.FScrollCount + oCMonthlyMaeipLedge.StarScrollPage - 1 %>
        		<% if i>oCMonthlyMaeipLedge.FTotalpage then Exit for %>
        		<% if CStr(page)=CStr(i) then %>
        		<font color="red">[<%= i %>]</font>
        		<% else %>
        		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
        		<% end if %>
        	<% next %>

        	<% if oCMonthlyMaeipLedge.HasNextScroll then %>
        		<a href="javascript:NextPage('<%= i %>')">[next]</a>
        	<% else %>
        		[next]
        	<% end if %>
		</td>
	</tr>

</table>



<%
set oCMonthlyMaeipLedge = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
