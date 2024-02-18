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
Const isOnlySys = FALSE
Dim isViewUser : isViewUser = FALSE ''(session("ssAdminPsn")="17")

dim yyyy1,mm1,isusing,sysorreal, research, newitem, vatyn, minusinc, bPriceGbn
dim mwgubun, buseo, itemgubun
dim purchasetype, swSppPrc

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
swSppPrc	= requestCheckvar(request("swSppPrc"),32)

if (sysorreal="") then sysorreal="sys"  ''real
if (isViewUser="") then sysorreal="sys"
if (isViewUser="") then bPriceGbn="P"
if (isViewUser="") then isusing=""

if (research="") then
	buseo = "3X"
    bPriceGbn="P"
	swSppPrc = "Y"
	mwgubun = "M"
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
ojaego.FRectIsUsing = isusing
ojaego.FRectGubun = sysorreal
ojaego.FRectNewItem = newitem
ojaego.FRectMwDiv    = mwgubun
ojaego.FRectVatYn    = vatyn
ojaego.FRectItemGubun = itemgubun
ojaego.FRectMinusInclude = minusinc
ojaego.FRectPurchaseType = purchasetype
ojaego.FRectTargetGbn = buseo
ojaego.FRectShopSuplyPrice    = swSppPrc

if (buseo="IT") then
    ojaego.FRectITSOnlyOrNot = "O"
else
    ojaego.FRectITSOnlyOrNot = "N"
end if

if (bPriceGbn="P") then
    ojaego.FRectIsFix = "on"
end if
ojaego.GetMonthlyJeagoSumWithPreMonth '' GetMonthlyJeagoSumNew ''


dim i
dim totno, totbuy, subTotno, subTotbuy '', totavgBuy, offtotavgBuy

dim totPreno, totPrebuy     , subPreno, subPrebuy
dim totIpno,totIpBuy        , subIpno, subIpBuy
dim totLossno, totLossBuy   , subLossno, subLossBuy


dim iURL
dim nBusiName
%>
<script type='text/javascript'>

function reActMonthSummary(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 월말 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=monthlystock&yyyymm=' + yyyymm,'reActMonthSummary','width=600,height=600');
	popwin.focus();
}
function reActdailySummary1(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm('일별입출고 STEP1 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=dailystock1&yyyymm=' + yyyymm,'reActMonthSummary','width=600,height=600');
	popwin.focus();
}
function reActdailySummary2(){

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm('일별입출고 STEP2 재고 내역을 재작성 하시겠습니까?')){ return; }

	var popwin = window.open('<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp?mode=dailystock2&yyyymm=' + yyyymm,'reActMonthSummary','width=600,height=600');
	popwin.focus();
}

</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %> 말일자 재고자산
			<!--
	        	&nbsp;&nbsp;
	        	<input type="radio" name="newitem" value="" <% if newitem="" then response.write "checked" %> >전체상품
	        	<input type="radio" name="newitem" value="new" <% if newitem="new" then response.write "checked" %> >신상품
	         -->
	        	&nbsp;&nbsp;|&nbsp;&nbsp;

	        	과세구분
	        	<input type="radio" name="vatyn" value="" <% if vatyn="" then response.write "checked" %> >전체
	        	<input type="radio" name="vatyn" value="Y" <% if vatyn="Y" then response.write "checked" %> >과세
	        	<input type="radio" name="vatyn" value="N" <% if vatyn="N" then response.write "checked" %> >면세
	        	&nbsp;&nbsp;
			    <input type="checkbox" name="swSppPrc" value="Y" <%= CHKIIF(swSppPrc="Y","checked","") %> >공급가로 표시
		</td>

		<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		<% IF Not (isOnlySys) THEN %>
		    <% if (Not isViewUser) then %>
			<font color="#CC3333">재고구분:</font>
        	<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >시스템재고
        	<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >실사재고
        	&nbsp;&nbsp;
        	<% end if %>
        <% END IF %>

        <% if (Not isViewUser) then %>
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

		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">

		<font color="#CC3333">마이너스구분:</font>
		<input type="radio" name="minusinc" value="" <%= CHKIIF(minusinc="","checked","") %> >마이너스재고 포함(전체)
		<input type="radio" name="minusinc" value="P" <%= CHKIIF(minusinc="P","checked","") %> >(+)재고만
	    <input type="radio" name="minusinc" value="M" <%= CHKIIF(minusinc="M","checked","") %> >마이너스재고 만
	    &nbsp;&nbsp;
	    <% if (Not isViewUser) then %>
	    <font color="#CC3333">매입가기준:</font>
	    <input type="radio" name="bPriceGbn" value="" <%= CHKIIF(bPriceGbn="","checked","") %>  >현재매입가
	    <input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >작성시매입가
	    <input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %> disabled >평균매입가
	    <% end if %>
	    </td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	    	<font color="#CC3333">부서구분:</font>
	        <% Call drawSelectBoxBuseoGubunWith3PL("buseo", buseo) %>
			&nbsp;
	    	<font color="#CC3333">상품구분:</font>
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			&nbsp;
			구매유형 : <% drawPartnerCommCodeBox True, "purchasetype", "purchasetype", purchasetype, "" %>
	    </td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

	<!--* <font color="red">오늘 작성한 입출고내역</font>은 제외됩니다.(야간에 일괄 반영됨, 상품별재고현황에서 입출고내역 전체 새로고침하면 반영됨)<br>-->
	* <font color="red">정산확정 후</font> 출고내역을 작성한 경우 재고자산에 반영안됨(서이사님에게 요청해야 반영됨-mwgubun)

<p>

<!-- 액션 시작 -->
<% ''if C_ADMIN_AUTH or (session("ssBctId") = "faxy") then %>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="일별입출고재작성STEP1" onclick="reActdailySummary1();" <% if (Left(DateAdd("m", -1, Now()), 7) > (yyyy1 + "-" + mm1)) then %>disabled<% end if %> ><!-- 문재 요청으로 disable -->
			<input type="button" class="button" value="일별입출고재작성STEP2" onclick="reActdailySummary2();" <% if (Left(DateAdd("m", -1, Now()), 7) > (yyyy1 + "-" + mm1)) then %>disabled<% end if %> ><!-- 문재 요청으로 disable -->
			<input type="button" class="button" value="재고자산재작성" onclick="reActMonthSummary();" <% if (Left(DateAdd("m", -1, Now()), 7) > (yyyy1 + "-" + mm1)) then %>disabled<% end if %> ><!-- 문재 요청으로 disable -->
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
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="4">상품구분</td>
        <td colspan="2">기초재고(월말일자)<br>A</td>
        <td colspan="2">당월매입(월)<br>B</td>
        <td colspan="2">기말재고(월말일자)<br>C</td>
        <td colspan="2">총매출원가<br>D=A+B-C</td>
        <td width="1" bgcolor="#FFFFFF"></td>
        <td colspan="2">재고LOSS<br>E</td>
        <td colspan="2">상품매출원가<br>F=A+B+E-C</td>
    </tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	    <td >부서</td>
	    <td >구분</td>
	    <td >코드구분</td>
    	<td >매입구분</td>
    	<td >수량</td>
    	<td >금액(매입가)</td>
    	<td >수량</td>
    	<td >금액(매입가)</td>
    	<td >수량</td>
    	<td >금액(매입가)</td>
    	<td >수량</td>
    	<td >금액(매입가)</td>
    	<td  bgcolor="#FFFFFF"></td>
    	<td >수량</td>
    	<td >금액(매입가)</td>
    	<td >수량</td>
    	<td >금액(매입가)</td>
    </tr>
    <% for i=0 to ojaego.FResultCount-1 %>
    <%
    if i<ojaego.FResultCount-1 then
        nBusiName= ojaego.FItemList(i+1).getBusiName
    else
        nBusiName=""
    end if

    if (ojaego.FItemList(i).getBusiName=nBusiName) then nBusiName=""
    if (i=ojaego.FResultCount-1) then nBusiName="L"

    totno   = totno + ojaego.FItemList(i).FTotCount
    totbuy  = totbuy + ojaego.FItemList(i).FTotBuySum

    totPreno = totPreno + ojaego.FItemList(i).FTotPreCount
    totPrebuy= totPrebuy + ojaego.FItemList(i).FTotPreBuySum

    totIpno  = totIpno + ojaego.FItemList(i).FTotIpCount
    totIpBuy = totIpBuy + ojaego.FItemList(i).FTotIpBuySum

    totLossno  = totLossno + ojaego.FItemList(i).FTotLossCount
    totLossBuy = totLossBuy + ojaego.FItemList(i).FTotLossBuySum

    subTotno    = subTotno + ojaego.FItemList(i).FTotCount
    subTotbuy   = subTotbuy + ojaego.FItemList(i).FTotBuySum

    subPreno    = subPreno + ojaego.FItemList(i).FTotPreCount
    subPrebuy   = subPrebuy + ojaego.FItemList(i).FTotPreBuySum
    subIpno     = subIpno + ojaego.FItemList(i).FTotIpCount
    subIpBuy    = subIpBuy + ojaego.FItemList(i).FTotIpBuySum
    subLossno   = subLossno + ojaego.FItemList(i).FTotLossCount
    subLossBuy  = subLossBuy + ojaego.FItemList(i).FTotLossBuySum


    iURL = "monthlystock_detail.asp?menupos="& menupos &"&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
    iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype&"&swSppPrc="&swSppPrc
    if Not(isOnlySys) THEN iURL=iURL&"&sysorreal="& sysorreal
    %>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%= ojaego.FItemList(i).getBusiName %></td>
        <td><a href="<%= iURL %>" target="_blank"><%= GetItemGubunName(ojaego.FItemList(i).Fitemgubun) %></a></td>
        <td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).Fitemgubun %></a></td>
    	<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotPreCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotPreBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotIpCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotIpBuySum,0) %></td>
        <td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getWongaCnt,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getWongaSum,0) %></td>
    	<td ></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotLossCount,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotLossBuySum,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaCnt,0) %></td>
    	<td align="right"><%= FormatNumber(ojaego.FItemList(i).getLossAssignedWongaSum,0) %></td>
    </tr>

    <% if (nBusiName<>"") then %>
    <tr align="center" bgcolor="#EEFFEE">

    	<td></td>
    	<td>소계</td>
    	<td></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(subPreno,0) %></td>
    	<td align="right" ><%= FormatNumber(subPrebuy,0) %></td>
    	<td align="right" ><%= FormatNumber(subIpno,0) %></td>
    	<td align="right" ><%= FormatNumber(subIpBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(subTotno,0) %></td>
    	<td align="right" ><%= FormatNumber(subTotbuy,0) %></td>
    	<td align="right" ><%= FormatNumber(subPreno+subIpno-subtotno,0) %></td>
    	<td align="right" ><%= FormatNumber(subPrebuy+subIpBuy-subtotbuy,0) %></td>
    	<td ></td>
    	<td align="right" ><%= FormatNumber(subLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(subLossBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(subPreno+subIpno-subtotno+subLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(subPrebuy+subIpBuy-subtotbuy+subLossBuy,0) %></td>
    </tr>
    <tr  bgcolor="#FFFFFF">
    	<td colspan="17"></td>
    </tr>
    <%
        subTotno=0
        subTotbuy=0

        subPreno   =0
        subPrebuy  =0
        subIpno    =0
        subIpBuy    =0
        subLossno   =0
        subLossBuy  =0

    %>
    <% end if %>
    <% next %>



    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
    	<td>총계</td>
    	<td></td>
    	<td></td>
    	<td align="right" ><%= FormatNumber(totPreno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totIpno,0) %></td>
    	<td align="right" ><%= FormatNumber(totIpBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totbuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totPreno+totIpno-totno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy+totIpBuy-totbuy,0) %></td>
    	<td ></td>
    	<td align="right" ><%= FormatNumber(totLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(totLossBuy,0) %></td>
    	<td align="right" ><%= FormatNumber(totPreno+totIpno-totno+totLossno,0) %></td>
    	<td align="right" ><%= FormatNumber(totPrebuy+totIpBuy-totbuy+totLossBuy,0) %></td>
    </tr>
</table>



<%
set ojaego = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
