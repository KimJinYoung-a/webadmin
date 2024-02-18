<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyInventoryCls.asp"-->
<%

dim research, i
dim yyyy1,mm1, yyyymm1, makerid, showsupply
dim stockPlace, shopid, stockGubun, showShopid
dim targetGbn, itemgubun, mwdiv
dim ArrList

research    = requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1       	= requestCheckvar(request("mm1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
stockGubun  = requestCheckvar(request("stockGubun"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsupply   = requestCheckvar(request("showsupply"),10)
shopid    	= requestCheckvar(request("shopid"),32)
itemgubun   = requestCheckvar(request("itemgubun"),10)
targetGbn   = requestCheckvar(request("targetGbn"),10)
mwdiv   = requestCheckvar(request("mwdiv"),10)
showShopid   = requestCheckvar(request("showShopid"),10)


dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if


yyyymm1 = yyyy1 + "-" + mm1


dim oCMonthlyInventory
set oCMonthlyInventory = new CMonthlyInventory

oCMonthlyInventory.FRectYYYYMM = yyyymm1
oCMonthlyInventory.FRectStockPlace = stockPlace
oCMonthlyInventory.FRectStockGubun = stockGubun
oCMonthlyInventory.FRectMakerid = makerid
oCMonthlyInventory.FRectBySupplyPrice = showsupply
oCMonthlyInventory.FRectShopid = shopid
oCMonthlyInventory.FRectItemgubun = itemgubun
oCMonthlyInventory.FRectTargetGbn = targetGbn
oCMonthlyInventory.FRectMwdiv = mwdiv
oCMonthlyInventory.FRectShowShopid = showShopid

oCMonthlyInventory.GetMonthlyInventorySUM

if oCMonthlyInventory.FTotalCount>0 then
	ArrList = oCMonthlyInventory.farrlist
end if

dim oitem

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
    realCall(yyyymm, 'makeStockEndStock');
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
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="" target="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">* 연/월 :</font> <% DrawYMBox yyyy1,mm1 %>
			&nbsp;&nbsp;
			<font color="#CC3333">* 브랜드 :</font> <%	drawSelectBoxDesignerWithName "makerid", makerid %>
            &nbsp;&nbsp;
            <font color="#CC3333">* 매장 :</font> <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showsupply" value="Y" <%= CHKIIF(showsupply="Y","checked","") %> disabled> 공급가로 표시
            &nbsp;&nbsp;
			<input type="checkbox" name="showShopid" value="Y" <%= CHKIIF(showShopid="Y", "checked", "") %> > 매장표시
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.target='';document.frm.action='';document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
            <font color="#CC3333">* 매입구분 :</font>
		    <select name="mwdiv" class="select">
		        <option value="" <%= CHKIIF(mwdiv="","selected" ,"") %> >전체</option>
        		<option value="M" <%= CHKIIF(mwdiv="M","selected" ,"") %> >매입</option>
        		<option value="W" <%= CHKIIF(mwdiv="W","selected" ,"") %> >위탁</option>
        	</select>
        	&nbsp;&nbsp;
		    <font color="#CC3333">* 재고위치 :</font>
		    <select name="stockPlace" class="select">
		        <option value="" <%= CHKIIF(stockPlace="","selected" ,"") %> >전체</option>
        		<option value="L" <%= CHKIIF(stockPlace="L","selected" ,"") %> >물류</option>
        		<option value="S" <%= CHKIIF(stockPlace="S","selected" ,"") %> >매장</option>
				<option value="E" <%= CHKIIF(stockPlace="E","selected" ,"") %> >기타</option>
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">* 재고구분 :</font>
        	<select name="stockGubun" class="select">
        	<option value="">전체
        	<option value="M" <%= CHKIIF(stockGubun="M","selected" ,"") %> >매입
        	<option value="W" <%= CHKIIF(stockGubun="W","selected" ,"") %> >위탁
        	<option value="T" <%= CHKIIF(stockGubun="T","selected" ,"") %> >3PL
        	</select>
        	&nbsp;&nbsp;
        	<font color="#CC3333">* 코드구분 :</font>
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			&nbsp;&nbsp;
			<font color="#CC3333">* 매입가기준 :</font>
			<input type="radio" name="priceGubun" value="V" checked> 평균매입가
	    </td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

* 3월 CS내역 사라짐 : 10-2867096-0012

<p />

<div style="float: right; margin-bottom: 5px;">
    <input type="button" value="재작성(<%= yyyy1 & "-" & mm1 %>)" onclick="jsRewrite('<%= yyyy1 & "-" & mm1 %>')" />
</div>

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="5">상품구분</td>
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
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>구분</td>
	    <td>코드<br />구분</td>
	    <td>매입<br />구분</td>
	    <td>재고<br>위치</td>
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
    </tr>
    <% if isarray(arrlist) then %>
    <% for i=0 to ubound(arrlist,2) %>
    <%
    set oitem = new CMonthlyInventoryItem
    Call oitem.SetValueByArray(arrlist, i)
    %>
    <tr align="right" bgcolor="#FFFFFF" >
        <td align="center"><%= oitem.GetShopDivBasic() %></td>
        <td align="center"><%= oitem.Fitemgubun %></td>
        <td align="center"><%= oitem.GetMwdivName %></td>
        <td align="center"><%= oitem.GetStockPlaceName %></td>
        <td align="center"><%= oitem.Fshopid %></td>
        <td><%= FormatNumber(oitem.FBeginingNo, 0) %></td>
        <td><%= FormatNumber(oitem.FBeginingSum, 0) %></td>
        <td><%= FormatNumber(oitem.FMaeipNo, 0) %></td>
        <td><%= FormatNumber(oitem.FMaeipSum, 0) %></td>
        <td><%= FormatNumber(oitem.FMoveNo, 0) %></td>
        <td><%= FormatNumber(oitem.FMoveSum, 0) %></td>
        <td><%= FormatNumber(oitem.FSellNo, 0) %></td>
        <td><%= FormatNumber(oitem.FSellSum, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgoOneNo, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgSOneum, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgoTwoNo, 0) %></td>
        <td><%= FormatNumber(oitem.FChulgoTwoSum, 0) %></td>
        <td><%= FormatNumber(oitem.FEtcChulgoNo, 0) %></td>
        <td><%= FormatNumber(oitem.FEtcChulgoSum, 0) %></td>
        <td><%= FormatNumber(oitem.FCsChulgoNo, 0) %></td>
        <td><%= FormatNumber(oitem.FCsChulgoSum, 0) %></td>
        <td><%= FormatNumber(oitem.getDiffNo, 0) %></td>
        <td><%= FormatNumber(oitem.getDiffSum, 0) %></td>
        <td><%= FormatNumber(oitem.FEndingNo, 0) %></td>
        <td><%= FormatNumber(oitem.FEndingNo, 0) %></td>
    </tr>
    <% next %>
    <% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
