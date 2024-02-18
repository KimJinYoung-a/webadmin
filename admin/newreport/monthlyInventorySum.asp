<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyInventoryCls.asp"-->
<%

dim research, i
dim yyyy1,mm1, yyyymm1, makerid, showsupply
dim stockPlace, shopid, stockGubun, showShopid, showMakerid, showItemid
dim targetGbn, itemgubun, mwdiv, page, itemid, hasOnly
dim ArrList, csDiffStr

research    = requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1       	= requestCheckvar(request("mm1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
stockGubun  = requestCheckvar(request("stockGubun"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsupply  = requestCheckvar(request("showsupply"),10)
shopid    	= requestCheckvar(request("shopid"),32)
itemgubun   = requestCheckvar(request("itemgubun"),10)
itemid   	= requestCheckvar(request("itemid"),32)
targetGbn   = requestCheckvar(request("targetGbn"),10)
mwdiv   	= requestCheckvar(request("mwdiv"),10)
hasOnly   	= requestCheckvar(request("hasOnly"),32)

showShopid   = requestCheckvar(request("showShopid"),10)
showMakerid   = requestCheckvar(request("showMakerid"),10)
showItemid   = requestCheckvar(request("showItemid"),10)

page   = requestCheckvar(request("page"),10)


dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


yyyymm1 = yyyy1 + "-" + mm1

if page = "" then
    page = "1"
end if

if research = "" then
    stockGubun = "TEN"
    mwdiv = "M"
end if



dim oCMonthlyInventory
set oCMonthlyInventory = new CMonthlyInventory

oCMonthlyInventory.FRectYYYYMM = yyyymm1
oCMonthlyInventory.FRectItemgubun = itemgubun

if (itemid <> "") then
    oCMonthlyInventory.FRectItemID = itemid

    oCMonthlyInventory.FRectShowShopid = "Y"
    oCMonthlyInventory.FRectShowMakerid = "Y"
    oCMonthlyInventory.FRectShowItemid = "Y"
    oCMonthlyInventory.FRectBySupplyPrice = showsupply
else
    oCMonthlyInventory.FRectStockPlace = stockPlace
    oCMonthlyInventory.FRectStockGubun = stockGubun
    oCMonthlyInventory.FRectMakerid = makerid
    oCMonthlyInventory.FRectBySupplyPrice = showsupply
    oCMonthlyInventory.FRectShopid = shopid

    oCMonthlyInventory.FRectTargetGbn = targetGbn
    oCMonthlyInventory.FRectMwdiv = mwdiv
    oCMonthlyInventory.FRectHasOnly = hasOnly

    oCMonthlyInventory.FRectShowShopid = showShopid
    oCMonthlyInventory.FRectShowMakerid = showMakerid
    oCMonthlyInventory.FRectShowItemid = showItemid
end if


oCMonthlyInventory.FCurrPage = page
oCMonthlyInventory.FPageSize = 1000

oCMonthlyInventory.GetMonthlyInventorySUM

if oCMonthlyInventory.FTotalCount>0 then
	ArrList = oCMonthlyInventory.farrlist
end if

oCMonthlyInventory.FRectMwdiv = mwdiv
csDiffStr = oCMonthlyInventory.GetMonthlyInventoryCSDiffList

dim oitem, dstart, dend

dim totBeginingNo
dim totBeginingSum
dim totMaeipNo
dim totMaeipSum
dim totMoveNo
dim totMoveSum
dim totSellNo
dim totSellSum
dim totChulgoOneNo
dim totChulgoOneSUM
dim totChulgoTwoNo
dim totChulgoTwoSum
dim totEtcChulgoNo
dim totEtcChulgoSum
dim totCsChulgoNo
dim totCsChulgoSum
dim totDiffNo
dim totDiffSum
dim totEndingNo
dim totEndingSum

%>
<script src="/js/jquery-1.7.1.min.js"></script>
<script>
function jsRewrite(yyyymm) {
    // 기초재고
    realCall(yyyymm, 'makeStockBeginStock');
    realCall(yyyymm, 'makeStockIpgo');
    realCall(yyyymm, 'makeStockMove');
    realCall(yyyymm, 'makeStockSell');
    realCall(yyyymm, 'makeStockSellOnGift');
    realCall(yyyymm, 'makeStockSellUpcheWitak');
    realCall(yyyymm, 'makeStockShopLoss');
    realCall(yyyymm, 'makeStockCsChulgo');
    realCall(yyyymm, 'makeWitakSell2Maeip');
    realCall(yyyymm, 'makeStockEndStock');
}

function jsExcItem(yyyymm, shopid, itemgubun, itemid, itemoption) {
    var url;
    var mode = 'excitem';
    var host = window.location.protocol + "//" + window.location.host + '/admin/newreport/monthlyInventorySum_process.asp?yyyymm=' + yyyymm + '&silent=Y';

    url = host + '&mode=' + mode + '&shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption;
    var data = '{}';

    if (confirm('재고자산 제외상품에 등록하시겠습니까?') != true) {
        return;
    }

    $.ajax({
        type : 'POST',
        url : url,
        data : data,
        async : false,
        dataType: 'html',
        contentType: 'application/x-www-form-urlencoded; charset=euc-kr',
        error:function(request, status, error) {
            alert("code:"+request.status+"\n"+"message:"+request.responseText+"\n"+"error:"+error);
        },
        success : function(data) {
            if (data.indexOf('{') > 0) {
                data = data.substring(data.indexOf('{'));
            }
            // alert(data);

            var obj = JSON.parse(data);
            if (obj.code == '000') {
                alert(obj.message);
            } else {
                alert(obj.message);
            }
        }
    });
}

function realCall(yyyymm, mode) {
    var url;
    var host = window.location.protocol + "//" + window.location.host + '/admin/newreport/monthlyInventorySum_process.asp?yyyymm=' + yyyymm + '&silent=Y';

    url = host + '&mode=' + mode;
    var data = '{}';

    $.ajax({
        type : 'POST',
        url : url,
        data : data,
        async : false,
        dataType: 'html',
        contentType: 'application/x-www-form-urlencoded; charset=euc-kr',
        error:function(request, status, error) {
            alert("code:"+request.status+"\n"+"message:"+request.responseText+"\n"+"error:"+error);
        },
        success : function(data) {
            if (data.indexOf('{') > 0) {
                data = data.substring(data.indexOf('{'));
            }
            // alert(data);

            var obj = JSON.parse(data);
            if (obj.code == '000') {
                alert(obj.message);
            } else {
                alert(obj.message);
            }
        }
    });
}

function jsPop(itemgubun, mwdiv, stockPlace) {
    var itemgubun2 = $('#frm').find('select[name="itemgubun"]').val();
    var mwdiv2 = $('#frm').find('select[name="mwdiv"]').val();
    var stockPlace2 = $('#frm').find('select[name="stockPlace"]').val();

    $('#frm').find('select[name="itemgubun"]').val(itemgubun);
    $('#frm').find('select[name="mwdiv"]').val(mwdiv);
    $('#frm').find('select[name="stockPlace"]').val(stockPlace);

    var params = $('#frm').serialize();
    var url = "monthlyInventorySum_Detail.asp?" + params;
    var popwin = window.open(url, "_blank","width=1800 height=800 scrollbars=yes resizable=yes status=yes");
	popwin.focus();

    $('#frm').find('select[name="itemgubun"]').val(itemgubun2);
    $('#frm').find('select[name="mwdiv"]').val(mwdiv2);
    $('#frm').find('select[name="stockPlace"]').val(stockPlace2);
}

function jsPopChulgoOne(yyyymm, shopid, itemid) {
    var yyyy, mm;
    yyyy = yyyymm.substring(0, 4);
    mm = yyyymm.substring(5);

    // month 는 0 부터 시작
    var d = new Date(yyyy*1, mm*1, 0);
    var dd = d.getDate();

    var popwin = window.open("/admin/newstorage/culgolist.asp?menupos=540&designer=" + shopid + "&itemid=" + itemid + "&chulgocheck=on&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + dd,"jsPopChulgoOne","width=1600,height=800 scrollbars=yes resizable=yes");
	popwin.focus();
}

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function jsSubmit() {
    var frm = document.frm;
    frm.submit();
}

function PopItemUpcheIpChulListOffLine(fromdate, todate, itemgubun, itemid, itemoption, ipchulflag, shopid) {
	var popwin = window.open('/common/pop_upcheipgolist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function jsPopEtcChulgo(fromdate, todate, shopid, shopdiv, stockPlace, itemgubun, itemid, itemoption) {
    if ((stockPlace == 'SL') || (stockPlace == 'SS')) {
        PopItemUpcheIpChulListOffLine(fromdate, todate, itemgubun, itemid, itemoption, 'S', shopid);
    } else if (stockPlace == 'O') {
        PopItemIpChulList(fromdate, todate, itemgubun, itemid, itemoption,'E');
    } else {
        alert('표시할 정보가 없습니다.');
    }
}

function popBuyItemListChulgo(ostr, itemgubun, itemid, itemoption) {
    if (ostr.length==7){
        var yyyy1   =   ostr.substr(0,4);
        var mm1     =   ostr.substr(5,2);
        var dd1     =   '01';

        var lastdate = new Date(yyyy1,mm1*1+1,0);
        var lastdate2 = new Date(yyyy1,mm1,0);

        var yyyy2   =   lastdate.getFullYear().toString(); //lastdate.getYear().toString();
        var mm2     =   lastdate.getMonth().toString();
        var dd2     =   lastdate2.getDate().toString();

        if (mm2.length<2) { mm2 = '0' + mm2 };
        if (dd2.length<2) { dd2 = '0' + dd2 };

    }else{
        var yyyy1   =   ostr.substr(0,4);
        var mm1     =   ostr.substr(5,2);
        var dd1     =   ostr.substr(8,2);

        var yyyy2   =   yyyy1;
        var mm2     =   mm1;
        var dd2     =   dd1;
    }

    var rectStr = '&yyyy1=' + yyyy1 + '&mm1=' + mm1 + '&dd1=' + dd1 + '&yyyy2=' + yyyy2 + '&mm2=' + mm2 + '&dd2=' + dd2;

	var popwin;
    if (itemgubun == '85') {
        popwin = window.open('/admin/ordermaster/onegiftitembuylist.asp?itemgubun=' + itemgubun+ '&itemid=' + itemid+ '&itemoption=' + itemoption+ '&itemstate=8&menupos=1527&datetype=beasong' + rectStr ,'popBuyItemList','width=1200,height=460,scrollbars=yes,resizable=yes');
    } else {
        popwin = window.open('/admin/ordermaster/oneitembuylist.asp?itemgubun=' + itemgubun+ '&itemid=' + itemid+ '&itemoption=' + itemoption+ '&itemstate=8&menupos=77&datetype=beasong' + rectStr ,'popBuyItemList','width=1200,height=460,scrollbars=yes,resizable=yes');
    }
	popwin.focus();
}

function PopItemSellList(fromdate, todate, shopid, shopdiv, stockPlace, itemgubun, itemid, itemoption) {
    if (itemid == '') {
        alert('상품코드가 없습니다.');
        return;
    }

    if (stockPlace == 'S') {
        PopItemSellListOffLine(fromdate, todate, itemgubun, itemid, itemoption, 'S', shopid);
    } else if (stockPlace == 'L') {
        popBuyItemListChulgo(fromdate.substring(0, 7), itemgubun, itemid, itemoption);
    } else {
        alert('표시할 정보가 없습니다.');
    }
}

function PopItemSellListOffLine(fromdate, todate, itemgubun, itemid, itemoption, ipchulflag, shopid) {
	var popwin = window.open('/common/pop_selllist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'PopItemSellListOffLine','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemStock(shopid, itemgubun, itemid, itemoption, barcode) {
    if (shopid == '') {
        var popwin = window.open("/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun=" + itemgubun + "&itemid=" + itemid + "&itemoption=" + itemoption,"PopItemStock","width=1200 height=600 scrollbars=yes resizable=yes");
	    popwin.focus();
    } else {
        var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?menupos=1075&shopid=' + shopid + '&barcode=' + barcode,'PopItemStock','width=1200,height=600,scrollbars=yes,resizable=yes')
	    popwin.focus();
    }
}

function PopItemUpcheIpChulListOffLine(fromdate,todate,itemgubun,itemid,itemoption, ipchulflag, shopid){
	var popwin = window.open('/common/pop_upcheipgolist_off.asp?fromdate=' + fromdate + '&todate=' + todate + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption + '&ipchulflag=' + ipchulflag + '&shopid=' + shopid,'poperritemlist','width=1000,height=600,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function PopItemMoveList(fromdate, todate, shopid, shopdiv, stockPlace, itemgubun, itemid, itemoption) {
    if (stockPlace == 'S') {
        PopItemUpcheIpChulListOffLine(fromdate, todate, itemgubun, itemid, itemoption, '', shopid);
    } else {
        alert('표시할 정보가 없습니다.');
    }
}

function popAccStockModiOne(itemgubun,itemid,itemoption){
	var popwin = window.open("/admin/newreport/pop_item_stock_Accsummary_edit.asp?itemgubun="+itemgubun+"&itemid=" + itemid + "&itemoption=" + itemoption,"popAccStockModiOne","width=1200 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>
<!-- 검색 시작 -->
<form id="frm" name="frm" method="get" action="" target="" style="margin: 0px;">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<input type="hidden" name="menupos" value="<%= menupos %>">
    <input type="hidden" name="page" value="">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">* 연/월 :</font> <% DrawYMBox yyyy1,mm1 %>
			&nbsp;&nbsp;
            <font color="#CC3333">* 상품코드 :</font> <input type="text" name="itemid" value="<%= itemid %>" size="8">
            &nbsp;&nbsp;
			<font color="#CC3333">* 브랜드 :</font> <%	drawSelectBoxDesignerWithName "makerid", makerid %>
            &nbsp;&nbsp;
            <font color="#CC3333">* 매장 :</font> <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showsupply" value="Y" <%= CHKIIF(showsupply="Y","checked","") %> > 공급가로 표시
            &nbsp;&nbsp;
			<input type="checkbox" name="showShopid" value="Y" <%= CHKIIF(showShopid="Y", "checked", "") %> > 매장표시
            &nbsp;&nbsp;
			<input type="checkbox" name="showMakerid" value="Y" <%= CHKIIF(showMakerid="Y", "checked", "") %> > 브랜드표시
            &nbsp;&nbsp;
			<input type="checkbox" name="showItemid" value="Y" <%= CHKIIF(showItemid="Y", "checked", "") %> > 상품코드표시
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="jsSubmit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
        	<font color="#CC3333">* 업체 :</font>
        	<select name="stockGubun" class="select">
        	<option value="">전체
        	<option value="TEN" <%= CHKIIF(stockGubun="TEN","selected" ,"") %> >텐바이텐
        	<option value="3PL" <%= CHKIIF(stockGubun="3PL","selected" ,"") %> >3PL
        	</select>
        	&nbsp;&nbsp;
		    <font color="#CC3333">* 재고위치 :</font>
		    <select name="stockPlace" class="select">
		        <option value="" <%= CHKIIF(stockPlace="","selected" ,"") %> >전체</option>
        		<option value="L" <%= CHKIIF(stockPlace="L","selected" ,"") %> >물류</option>
        		<option value="S" <%= CHKIIF(stockPlace="S","selected" ,"") %> >매장</option>
                <option value="J" <%= CHKIIF(stockPlace="J","selected" ,"") %> >정산</option>
				<option value="F" <%= CHKIIF(stockPlace="F","selected" ,"") %> >가맹</option>
                <option value="W" <%= CHKIIF(stockPlace="W","selected" ,"") %> >도매</option>
                <option value="R" <%= CHKIIF(stockPlace="R","selected" ,"") %> >랜탈</option>
                <option value="N" <%= CHKIIF(stockPlace="N","selected" ,"") %> >기타(사은품)</option>
                <option value="SS" <%= CHKIIF(stockPlace="SS","selected" ,"") %> >기타(매장샘플)</option>
                <option value="SL" <%= CHKIIF(stockPlace="SL","selected" ,"") %> >기타(매장로스)</option>
                <option value="O" <%= CHKIIF(stockPlace="O","selected" ,"") %> >기타(나머지)</option>
        	</select>
        	&nbsp;&nbsp;
            <font color="#CC3333">* 매입구분 :</font>
		    <select name="mwdiv" class="select">
		        <option value="" <%= CHKIIF(mwdiv="","selected" ,"") %> >전체</option>
        		<option value="M" <%= CHKIIF(mwdiv="M","selected" ,"") %> >매입</option>
        		<option value="W" <%= CHKIIF(mwdiv="W","selected" ,"") %> >위탁</option>
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">* 코드구분 :</font>
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			&nbsp;&nbsp;
			<font color="#CC3333">* 매입가기준 :</font>
			<input type="radio" name="priceGubun" value="V" checked> 평균매입가
            &nbsp;&nbsp;
            <select class="select" name="hasOnly">
                <option></option>
                <option value="diff" <%= CHKIIF(hasOnly="diff","selected" ,"") %>>오차</option>
                <option value="MoveNo" <%= CHKIIF(hasOnly="MoveNo","selected" ,"") %>>이동수량</option>
                <option value="CsChulgoNo" <%= CHKIIF(hasOnly="CsChulgoNo","selected" ,"") %>>CS수량</option>
                <option value="avgPrcZero" <%= CHKIIF(hasOnly="avgPrcZero","selected" ,"") %>>평균매입가 0원</option>
            </select>
            있는 상품만 표시
	    </td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->

<p />

* 상품코드를 입력하면 물류/매장 등 전체내역이 표시됩니다.<br />
* 공급가 계산<br />
&nbsp; - 당월매입 : 월매입총액에 대해 공급가 계산<br />
&nbsp; - 나머지 : 평균매입가에 공급가 계산 후 수량을 곱함<br />
* CS완료 후 삭제내역 : <%= csDiffStr %>
<p />

<div style="float: right; margin-bottom: 5px;">
    <input type="button" value="재작성(<%= yyyy1 & "-" & mm1 %>)" onclick="jsRewrite('<%= yyyy1 & "-" & mm1 %>')" />
</div>

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="6">상품구분</td>
        <% if showMakerid <> "" then %>
        <td rowspan="2">브랜드</td>
        <% end if %>
        <% if showItemid <> "" then %>
        <td rowspan="2">상품코드</td>
        <% end if %>
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
        <td>비고</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>업체</td>
        <td>재고<br>위치</td>
        <td>구분</td>
	    <td>코드<br />구분</td>
	    <td>매입<br />구분</td>
		<td>샵아이디</td>
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
        <td>비고</td>
    </tr>
    <% if isarray(arrlist) then %>
    <% for i=0 to ubound(arrlist,2) %>
    <%
    set oitem = new CMonthlyInventoryItem
    Call oitem.SetValueByArray(arrlist, i)

    dstart = yyyymm1 + "-01"
    dend = Left(DateAdd("m", 1, dstart), 7)+"-01"
    dend = Left(DateAdd("d", -1, dend), 10)

    totBeginingNo = totBeginingNo + oitem.FBeginingNo
    totBeginingSum = totBeginingSum + oitem.FBeginingSum
    totMaeipNo = totMaeipNo + oitem.FMaeipNo
    totMaeipSum = totMaeipSum + oitem.FMaeipSum
    totMoveNo = totMoveNo + oitem.FMoveNo
    totMoveSum = totMoveSum + oitem.FMoveSum
    totSellNo = totSellNo + oitem.FSellNo
    totSellSum = totSellSum + oitem.FSellSum
    totChulgoOneNo = totChulgoOneNo + oitem.FChulgoOneNo
    totChulgoOneSUM = totChulgoOneSUM + oitem.FChulgoOneSUM
    totChulgoTwoNo = totChulgoTwoNo + oitem.FChulgoTwoNo
    totChulgoTwoSum = totChulgoTwoSum + oitem.FChulgoTwoSum
    totEtcChulgoNo = totEtcChulgoNo + oitem.FEtcChulgoNo
    totEtcChulgoSum = totEtcChulgoSum + oitem.FEtcChulgoSum
    totCsChulgoNo = totCsChulgoNo + oitem.FCsChulgoNo
    totCsChulgoSum = totCsChulgoSum + oitem.FCsChulgoSum
    totDiffNo = totDiffNo + oitem.getDiffNo
    totDiffSum = totDiffSum + oitem.getDiffSum
    totEndingNo = totEndingNo + oitem.FEndingNo
    totEndingSum = totEndingSum + oitem.FEndingSum
    %>
    <tr align="right" bgcolor="#FFFFFF" >
        <td align="center"><%= oitem.GetStockGubunName() %></td>
        <td align="center"><%= oitem.FstockPlace %></td>
        <td align="center"><%= oitem.GetShopDivBasic() %></td>
        <td align="center"><%= oitem.Fitemgubun %></td>
        <td align="center"><%= oitem.GetMwdivName %></td>
        <td align="center"><%= oitem.Fshopid %></td>
        <% if showMakerid <> "" then %>
        <td align="left"><%= oitem.Fmakerid %></td>
        <% end if %>
        <% if showItemid <> "" then %>
        <td align="left"><a href="javascript:PopItemStock('<%= oitem.Fshopid %>', '<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>', '<%= BF_MakeTenBarcode(oitem.Fitemgubun, oitem.Fitemid, oitem.Fitemoption) %>')"><%= oitem.Fitemgubun %>-<%= oitem.Fitemid %>-<%= oitem.Fitemoption %></a></td>
        <% end if %>
        <td><%= FormatNumber(oitem.FBeginingNo, 0) %></td>
        <td><%= FormatNumber(oitem.FBeginingSum, 0) %></td>
        <td><%= FormatNumber(oitem.FMaeipNo, 0) %></td>
        <td><%= FormatNumber(oitem.FMaeipSum, 0) %></td>
        <td><a href="javascript:PopItemMoveList('<%= dstart %>', '<%= dend %>', '<%= oitem.Fshopid %>', '<%= oitem.GetShopDivBasic() %>', '<%= oitem.FstockPlace %>', '<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>')"><%= FormatNumber(oitem.FMoveNo, 0) %></a></td>
        <td><a href="javascript:PopItemMoveList('<%= dstart %>', '<%= dend %>', '<%= oitem.Fshopid %>', '<%= oitem.GetShopDivBasic() %>', '<%= oitem.FstockPlace %>', '<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>')"><%= FormatNumber(oitem.FMoveSum, 0) %></a></td>
        <td><a href="javascript:PopItemSellList('<%= dstart %>', '<%= dend %>', '<%= oitem.Fshopid %>', '<%= oitem.GetShopDivBasic() %>', '<%= oitem.FstockPlace %>', '<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>')"><%= FormatNumber(oitem.FSellNo, 0) %></a></td>
        <td><a href="javascript:PopItemSellList('<%= dstart %>', '<%= dend %>', '<%= oitem.Fshopid %>', '<%= oitem.GetShopDivBasic() %>', '<%= oitem.FstockPlace %>', '<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>')"><%= FormatNumber(oitem.FSellSum, 0) %></a></td>
        <td><a href="javascript:jsPopChulgoOne('<%= yyyymm1 %>', '<%= oitem.Fshopid %>', '<%= oitem.Fitemid %>')"><%= FormatNumber(oitem.FChulgoOneNo, 0) %></a></td>
        <td><a href="javascript:jsPopChulgoOne('<%= yyyymm1 %>', '<%= oitem.Fshopid %>', '<%= oitem.Fitemid %>')"><%= FormatNumber(oitem.FChulgoOneSUM, 0) %></a></td>
        <td><%= FormatNumber(oitem.FChulgoTwoNo, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgoTwoSum, 0) %></td>
        <td><a href="javascript:jsPopEtcChulgo('<%= dstart %>', '<%= dend %>', '<%= oitem.Fshopid %>', '<%= oitem.GetShopDivBasic() %>', '<%= oitem.FstockPlace %>', '<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>')"><%= FormatNumber(oitem.FEtcChulgoNo, 0) %></a></td>
        <td><a href="javascript:jsPopEtcChulgo('<%= dstart %>', '<%= dend %>', '<%= oitem.Fshopid %>', '<%= oitem.GetShopDivBasic() %>', '<%= oitem.FstockPlace %>', '<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>')"><%= FormatNumber(oitem.FEtcChulgoSum, 0) %></a></td>
        <td><%= FormatNumber(oitem.FCsChulgoNo, 0) %></td>
        <td><%= FormatNumber(oitem.FCsChulgoSum, 0) %></td>
        <td><%= FormatNumber(oitem.getDiffNo, 0) %></td>
        <td><%= FormatNumber(oitem.getDiffSum, 0) %></td>
        <td><%= FormatNumber(oitem.FEndingNo, 0) %></td>
        <td><%= FormatNumber(oitem.FEndingSum, 0) %></td>
        <td align="center">
            <img src="/images/icon_arrow_link.gif" style="cursor:pointer" onClick="popAccStockModiOne('<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>')">
            <% if C_ADMIN_AUTH then %>
            <% if (yyyymm1 <> "") and (oitem.Fitemgubun <> "") and (oitem.Fitemid <> "") and (oitem.Fitemoption <> "") then %>
            <% if (oitem.Fshopid <> "") or (oitem.FstockPlace = "L") or (oitem.FstockPlace = "O") or (oitem.FstockPlace = "N") or (oitem.FstockPlace = "R") then %>
            <a href="javascript:jsExcItem('<%= yyyymm1 %>', '<%= oitem.Fshopid %>', '<%= oitem.Fitemgubun %>', '<%= oitem.Fitemid %>', '<%= oitem.Fitemoption %>')">X</a>
            <% end if %>
            <% end if %>
            <% end if %>
        </td>
    </tr>
    <% next %>
    <tr align="right" bgcolor="#FFFFFF" >
        <td align="center" colspan="6"></td>
        <% if showMakerid <> "" then %>
        <td align="left"></td>
        <% end if %>
        <% if showItemid <> "" then %>
        <td align="left"></td>
        <% end if %>
        <td><%= FormatNumber(totBeginingNo, 0) %></td>
        <td><%= FormatNumber(totBeginingSum, 0) %></td>
        <td><%= FormatNumber(totMaeipNo, 0) %></td>
        <td><%= FormatNumber(totMaeipSum, 0) %></td>
        <td><%= FormatNumber(totMoveNo, 0) %></td>
        <td><%= FormatNumber(totMoveSum, 0) %></td>
        <td><%= FormatNumber(totSellNo, 0) %></td>
        <td><%= FormatNumber(totSellSum, 0) %></td>
        <td><%= FormatNumber(totChulgoOneNo, 0) %></td>
        <td><%= FormatNumber(totChulgoOneSUM, 0) %></td>
        <td><%= FormatNumber(totChulgoTwoNo, 0) %></td>
        <td><%= FormatNumber(totChulgoTwoSum, 0) %></td>
        <td><%= FormatNumber(totEtcChulgoNo, 0) %></td>
        <td><%= FormatNumber(totEtcChulgoSum, 0) %></td>
        <td><%= FormatNumber(totCsChulgoNo, 0) %></td>
        <td><%= FormatNumber(totCsChulgoSum, 0) %></td>
        <td><%= FormatNumber(totDiffNo, 0) %></td>
        <td><%= FormatNumber(totDiffSum, 0) %></td>
        <td><%= FormatNumber(totEndingNo, 0) %></td>
        <td><%= FormatNumber(totEndingSum, 0) %></td>
        <td></td>
    </tr>
    <% end if %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="<%= 27 + CHKIIF(showMakerid<>"", 1, 0) + CHKIIF(showItemid<>"", 1, 0) %>" align="center">
		<% if oCMonthlyInventory.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCMonthlyInventory.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCMonthlyInventory.StartScrollPage to oCMonthlyInventory.FScrollCount + oCMonthlyInventory.StartScrollPage - 1 %>
			<% if i>oCMonthlyInventory.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCMonthlyInventory.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>

</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
