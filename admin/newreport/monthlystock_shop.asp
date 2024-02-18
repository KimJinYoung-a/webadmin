<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%
Const isShowIpgoPrice = FALSE
Dim isShowSysWithReal : isShowSysWithReal = FALSE    '''시스템재고/실사재고 같이표시
Dim isViewUser : isViewUser = FALSE ''(session("ssAdminPsn")="17")

dim yyyy1,mm1,isusing,sysorreal, research, shopid, showminus
dim mwgubun, vatyn, showSupplyPrice, buseo
dim showminusOnly
dim etcjungsantype

yyyy1     = RequestCheckVar(request("yyyy1"),10)
mm1       = RequestCheckVar(request("mm1"),10)
isusing   = RequestCheckVar(request("isusing"),10)
sysorreal = RequestCheckVar(request("sysorreal"),10)
research  = RequestCheckVar(request("research"),10)
shopid    = RequestCheckVar(request("shopid"),32)
mwgubun   = RequestCheckVar(request("mwgubun"),10)
showminus   		= RequestCheckVar(request("showminus"),32)
vatyn       		= requestCheckvar(request("vatyn"),10)
showSupplyPrice 	= requestCheckvar(request("showSupplyPrice"),10)
buseo       		= requestCheckvar(request("buseo"),10)
showminusOnly       = requestCheckvar(request("showminusOnly"),10)
etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)

if (sysorreal="") then sysorreal="sys" ''real
if (isViewUser) then showminus=""
if (isViewUser) then showminusOnly=""
if (isViewUser) then sysorreal="sys"
if (isViewUser) then isusing=""

if (research="") and (showminus="") then showminus="on"
if (research="") and (mwgubun="") then mwgubun="M"
if (research="") and (etcjungsantype="") then etcjungsantype="41" ''직영+판매분
if (research="") and (buseo="") then buseo="3X" ''3pl제외
if (research="") and (showSupplyPrice="") then showSupplyPrice="Y"

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

dim oshopjaego
set oshopjaego = new CMonthlyStock
oshopjaego.FRectYYYYMM   = yyyy1 + "-" + mm1
oshopjaego.FRectYYYYMMDD = yyyy1 + "-" + mm1 + "-01"
oshopjaego.FRectIsUsing = isusing
oshopjaego.FRectGubun = sysorreal
oshopjaego.FRectShopid = shopid
oshopjaego.FRectMwDiv    = mwgubun
oshopjaego.FRectShowMinus = showminus
oshopjaego.FRectShowMinusOnly = showminusOnly
oshopjaego.FRectVatYn    = vatyn
oshopjaego.FRectShopSuplyPrice    = showSupplyPrice
oshopjaego.FRectTargetGbn = buseo
oshopjaego.FRectetcjungsantype = etcjungsantype

IF (isShowSysWithReal) then
    oshopjaego.FRectGubun = "sys"
    oshopjaego.GetShopMonthlyJeagoSumSysWithReal
ELSE
    oshopjaego.GetShopMonthlyJeagoSumNew
END IF

dim i
dim totno, totbuy, totsell, totavgBuy, offtotavgBuy
dim offtotno, offtotbuy, totshopBuy, offtotsell
dim totRealno, totRealbuy, totRealsell

dim iURL

%>
<script type='text/javascript'>

function reActdailySummary1(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm('일별입출고 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopdailystock1&yyyymm=' + yyyymm,'reActMonthSummary1','width=600,height=600');
	popwin.focus();
}
function reActMonthSummary(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm('월말 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthlystock&yyyymm=' + yyyymm,'reActMonthSummary','width=600,height=600');
	popwin.focus();
}

function reActMonthSummary10(){
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly10&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_1() {
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역(월별재고서머리)을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly101&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_2() {
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역(대학로매장)을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly102&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_3() {
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역(대학로 외 직영)을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly103&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_4() {
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역(띵소+출고위탁)을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly104&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary10_5() {
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역(제주 등)을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly105&yyyymm=' + yyyymm,'reActMonthSummary10','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary11(){
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly11&yyyymm=' + yyyymm,'reActMonthSummary11','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary20(){
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly20&yyyymm=' + yyyymm,'reActMonthSummary20','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary21(){
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly21&yyyymm=' + yyyymm,'reActMonthSummary21','width=100,height=100');
	popwin.focus();
}

function reActMonthSummary30(){
    //alert('수정중..');
    //return;

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=shopmonthly30&yyyymm=' + yyyymm,'reActMonthSummary30','width=100,height=100');
	popwin.focus();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 말일자 재고자산
			&nbsp;&nbsp;
			매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %> &nbsp;&nbsp;
			&nbsp;&nbsp;
			<font color="#CC3333">부서구분:</font>
	        <% Call drawSelectBoxBuseoGubunWith3PL("buseo", buseo) %>
        	&nbsp;&nbsp;
        	<font color="#CC3333">정산방법:</font>
        	<% 'drawPartnerCommCodeBox true,"etcjungsantype","etcjungsantype",etcjungsantype,"" %>
        	<select class="select" name="etcjungsantype"  >
            <option value="">-선택-</option>
            <option value="1" <%=CHKIIF(etcjungsantype="1","selected","")%> >판매분정산</option>
            <option value="2" <%=CHKIIF(etcjungsantype="2","selected","")%> >출고분정산</option>
            <option value="3" <%=CHKIIF(etcjungsantype="3","selected","")%> >가맹점정산</option>
            <option value="4" <%=CHKIIF(etcjungsantype="4","selected","")%> >직영점정산</option>
            <option value="41" <%=CHKIIF(etcjungsantype="41","selected","")%> >직영점+판매분정산</option>
            </select>

		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <% IF NOt (isShowSysWithReal) THEN %>
		    <% if (Not isViewUser) then %>
			<font color="#CC3333">재고구분:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >시스템재고
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >실사재고
        	&nbsp;&nbsp;
        	<% end if %>
        	<% end if %>

        	<% if (Not isViewUser) then %>
        	<font color="#CC3333">상품사용구분:</font>
        	<input type="radio" name="isusing" value="" <% if isusing="" then response.write "checked" %> >전체
        	<input type="radio" name="isusing" value="Y" <% if isusing="Y" then response.write "checked" %> >사용함
        	<input type="radio" name="isusing" value="N" <% if isusing="N" then response.write "checked" %> >사용안함
        	&nbsp;&nbsp;

        	<% end if %>

        	<font color="#CC3333">과세구분</font>
        	<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >전체
        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >과세
        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >면세
        	&nbsp;&nbsp;
			<input type="checkbox" name="showSupplyPrice" value="Y" <%= CHKIIF(showSupplyPrice="Y","checked","") %> >공급가로 표시
        	<br>
        	<font color="#CC3333">매입구분:</font>
        	<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >전체
        	<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >매입(매장매입+출고매입+ITS출고위탁)
        	<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >위탁(위탁판매+업체위탁+출고위탁)
        	<input type="radio" name="mwgubun" value="C" <% if mwgubun="C" then response.write "checked" %> >출고위탁
        	<!-- <input type="radio" name="mwgubun" value="U" <% if mwgubun="U" then response.write "checked" %> >업체 -->
        	<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >미지정
        	<% if (Not isViewUser) then %>
        	<br>
        	<input type="checkbox" name="showminus" <%= CHKIIF(showminus="on","checked","") %> >마이너스재고 포함
			<input type="checkbox" name="showminusOnly" <%= CHKIIF(showminusOnly="on","checked","") %> >마이너스재고만
        	<% end if %>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<br>

* <font color="red">오프라인 상품정보</font>가 없는 경우 표시되지 않습니다.<br />
* 매장 <font color="red">재고자산 매입구분</font>은 월말 정산내역이 작성될 때 결정됩니다.

<!-- 액션 시작 -->
<% ''if C_ADMIN_AUTH or (session("ssBctId") = "faxy") then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="일별입출고재작성" onclick="reActdailySummary1();" >
			<input type="button" class="button" value="재고자산재작성" onclick="reActMonthSummary();" >
			&nbsp;&nbsp;
			<!--
			<input type="button" class="button" value="재작성 1단계" onclick="reActMonthSummary10();">
			&nbsp;
			-->
			<input type="button" class="button" value="재작성 1-1단계" onclick="reActMonthSummary10_1();">
			<input type="button" class="button" value="재작성 1-2단계" onclick="reActMonthSummary10_2();">
			<input type="button" class="button" value="재작성 1-3단계" onclick="reActMonthSummary10_3();">
			<input type="button" class="button" value="재작성 1-4단계" onclick="reActMonthSummary10_4();">
			<input type="button" class="button" value="재작성 1-5단계" onclick="reActMonthSummary10_5();">
			<input type="button" class="button" value="재작성 2단계" onclick="reActMonthSummary11();">
			<input type="button" class="button" value="재작성 3-1단계" onclick="reActMonthSummary20();">
			<input type="button" class="button" value="재작성 3-2단계" onclick="reActMonthSummary21();">
			<input type="button" class="button" value="재작성 4단계" onclick="reActMonthSummary30();">
		</td>
		<td align="right">
		</td>
	</tr>
</table>

<p>
<% ''end if %>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td rowspan="2">매장</td>
        <td rowspan="2">정산방법</td>
    	<td width="110" rowspan="2">매입구분</td>
    	<td colspan="3">시스템재고</td>
    	<td width="39" rowspan="2">오차</td>
    	<td colspan="3">실사재고</td>
    	<td  width="90" rowspan="2">브랜드별<br>재고자산</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="110">재고수량</td>
    	<td width="110">소비자가*수량</td>
    	<td width="110">매입가*수량<br>(본사 매입가)</td>
    	<td width="110">재고수량</td>
    	<td width="110">소비자가*수량</td>
    	<td width="110">매입가*수량<br>(본사 매입가)</td>
    </tr>
	<% else %>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td width="110">부서</td>
	    <td width="110">매장ID</td>
	    <td >매장</td>
	    <td >정산방법</td>
    	<td  width="110">매입구분</td>
    	<td width="110">총재고수량</td>
    	<td width="110">소비자가*수량</td>
    	<!-- td width="110">평균마진</td -->
    	<td width="110">매입가*수량<br>(본사 매입가)</td>
    	<td width="110">공급가*수량<br>(매장 매입가)</td>
    	<% IF(isShowIpgoPrice)THEN %><td width="110">실매입가*수량</td><% END IF %>
    	<td  width="90">브랜드별<br>재고자산</td>
    	<!-- td  width="90">브랜드별<br>재고회전율</td -->
    </tr>
    <% end if %>

    <% for i=0 to oShopJaego.FResultCount-1 %>
    <% if (oShopJaego.FItemList(i).FMaeIpGubun="Z") and (oShopJaego.FItemList(i).FTotCount=0) then %>

    <% else %>
    <% if (TRUE) or oShopJaego.FItemList(i).FMaeIpGubun<>"Z" then %>
    <%
    totno   = totno + oShopJaego.FItemList(i).FTotCount
    totbuy  = totbuy + CCur(oShopJaego.FItemList(i).FTotBuySum)
    totshopBuy  = totshopBuy + CCur(oShopJaego.FItemList(i).FTotShopBuySum)
    totsell = totsell + CCur(oShopJaego.FItemList(i).FTotSellSum)

    if Not IsNULL(oShopJaego.FItemList(i).FavgIpgoPriceSum) THEN
       totavgBuy = totavgBuy + oShopJaego.FItemList(i).FavgIpgoPriceSum
    end if

    if (isShowSysWithReal) then
        totRealno   = totRealno + oShopJaego.FItemList(i).FTotRealCount
        totRealbuy  = totRealbuy + oShopJaego.FItemList(i).FTotRealBuySum
        totRealsell = totRealsell + oShopJaego.FItemList(i).FTotRealSellSum
    end if

    iURL = "monthlystockShop_detail.asp?menupos="& menupos &"&mwgubun="& oShopJaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1&"&mm1="& mm1 &"&isusing="& isusing &"&shopid="&oShopJaego.FItemList(i).FShopID&"&showminus="&showminus&"&showminusOnly="&showminusOnly&"&buseo="&oShopJaego.FItemList(i).FtargetGbn&"&vatyn="&vatyn&"&showSupplyPrice="&showSupplyPrice
    if Not(isShowSysWithReal) THEN iURL=iURL&"&sysorreal="& sysorreal
    %>
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).getBusiName %></a></td>
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).FShopName %></a></td>
    	<td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).getMaeipGubunName %></a></td>
    	<td><%= oShopJaego.FItemList(i).getEtcJungsanTypeName %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotBuySum,0) %></td>
    	<td align="center"><%= FormatNumber(oShopJaego.FItemList(i).FTotRealCount-oShopJaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotRealCount,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotRealSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotRealBuySum,0) %></td>
    	<td align="center"><a href="<%= iURL %>" target="_blank"><img src="/images/icon_search.jpg" width="16" border="0"></a></td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).getBusiName %></a></td>
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).FShopID %></a></td>
        <td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).FShopName %></a></td>
        <td><%= oShopJaego.FItemList(i).getEtcJungsanTypeName %></td>
    	<td><a href="<%= iURL %>" target="_blank"><%= oShopJaego.FItemList(i).getMaeipGubunName %></a></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotSellSum,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(oShopJaego.FItemList(i).FTotShopBuySum,0) %></td>
    	<% IF(isShowIpgoPrice)THEN %>
    	<td align="right">
        	<% If IsNULL(oShopJaego.FItemList(i).FavgIpgoPriceSum) then %>
        	-
        	<% else %>
        	<%= FormatNumber(oShopJaego.FItemList(i).FavgIpgoPriceSum,0) %>
        	<% end if %>
    	</td><% END IF %>
    	<td align="center"><a href="<%= iURL %>" target="_blank"><img src="/images/icon_search.jpg" width="16" border="0"></a></td>
    	<!--td align="center"><a href="javascript:alert('준비중.');" target="_blank"><img src="/images/icon_search.jpg" width="16" border="0"></a></td -->
    </tr>
    <% end if %>
    <% end if %>
    <% end if %>
    <% next %>
    <% if (isShowSysWithReal) then %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td >총계</td>
    	<td></td>
    	<td></td>
        <td></td>
        <td></td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td align="center" ><%= FormatNumber(totRealno-totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealno,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealsell,0) %></td>
    	<td align="right" ><%= FormatNumber(totRealbuy,0) %></td>
    	<td align="center"></td>
    </tr>
    <% else %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td >총계</td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totsell,0) %></td>
    	<!-- td></td -->
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totshopBuy,0) %></td>
        <% IF(isShowIpgoPrice)THEN %><td align="right"><%= FormatNumber(totavgBuy,0) %></td><% END IF %>
    	<td align="center"></td>
    	<!--td align="center"><a href="avascript:alert('준비중.');" target="_blank"><img src="/images/icon_search.jpg" width="16" border="0"></a></td-->
    </tr>
    <% end if %>
</table>



<%
set oShopJaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
