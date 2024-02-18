<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 재고월령fix
' History : 이상구 생성
'			2023.08.04 한용민 수정(상품구분별로 그룹지어지게 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls_2.asp"-->
<%
Dim isViewUser, yyyy1,mm1,isusing,sysorreal, research, newitem, minusinc, bPriceGbn, vatyn, mwgubun, buseo
dim itemgubun, mygubun, purchasetype, stplace, shopid, swSppPrc, etcjungsantype, nowdate, ojaego, i
dim subTotBuySum1, subTotBuySum2, subTotBuySum3, subTotBuySum4, subTotBuySum5, subTotBuySum6, subTotBuySum7
dim subTotBuySum8, subTotBuySum11, subTotBuySum12, subTotBuySum13, subTotBuySum14, subTotBuySum, subTotOverValueSum
dim sub_totStockNo, totBuySum1, totBuySum2, totBuySum3, totBuySum4, totBuySum5, totBuySum6, totBuySum7, totBuySum8
dim totBuySum11, totBuySum12, totBuySum13, totBuySum14, totBuySum, totOverValueSum, tot_totStockNo
dim totno, totbuy, subTotno, subTotbuy, totPreno, totPrebuy , subPreno, subPrebuy'', totavgBuy, offtotavgBuy
dim totIpno,totIpBuy, subIpno, subIpBuy, totLossno, totLossBuy, subLossno, subLossBuy, iURL
	yyyy1       = requestCheckvar(request("yyyy1"),10)
	mm1         = requestCheckvar(request("mm1"),10)
	isusing     = requestCheckvar(request("isusing"),10)
	sysorreal   = requestCheckvar(request("sysorreal"),10)
	research    = requestCheckvar(request("research"),10)
	newitem     = requestCheckvar(request("newitem"),10)
	mwgubun     = requestCheckvar(request("mwgubun"),10)
	mygubun     = requestCheckvar(request("mygubun"),10)
	minusinc   	= requestCheckvar(request("minusinc"),10)
	bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
	buseo       = requestCheckvar(request("buseo"),10)
	itemgubun   = requestCheckvar(request("itemgubun"),10)
	purchasetype   	= requestCheckvar(request("purchasetype"),10)
	vatyn       	= requestCheckvar(request("vatyn"),10)
	stplace       	= requestCheckvar(request("stplace"),10)
	shopid       	= requestCheckvar(request("shopid"),32)
	swSppPrc	= requestCheckvar(request("swSppPrc"),32)
	etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)

isViewUser = FALSE ''(session("ssAdminPsn")="17")
if (sysorreal="") then sysorreal="sys"  ''real
if (isViewUser="") then sysorreal="sys"
if (isViewUser="") then bPriceGbn="P"
if (isViewUser="") then isusing=""
if (research="") and (etcjungsantype="") then etcjungsantype="41" ''직영+판매분

if (research="") then
    bPriceGbn="V"
	buseo = "3X"
	mwgubun = "M"
	swSppPrc = "Y"
end if

if (stplace="") then
    stplace="L"
end if

if (mygubun = "") then
	mygubun = "M"
end if

if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

set ojaego = new CMonthlyMaeipLedge
	ojaego.FRectYYYYMM = yyyy1 + "-" + mm1
	ojaego.FRectGubun = "sys"						'//sysorreal
	ojaego.FRectMwDiv    = mwgubun
	ojaego.FRectItemGubun = itemgubun
	ojaego.FRectTargetGbn = buseo
	ojaego.FRectVatYn    = vatyn
	ojaego.FRectShopID    = shopid
	ojaego.FRectShopSuplyPrice    = swSppPrc
	ojaego.FRectetcjungsantype = etcjungsantype
	ojaego.FRectPriceGubun = bPriceGbn

	if (stplace = "L") then
		ojaego.GetJeagoOverValueSum
	else
		ojaego.FRectLastIpgoGBN = stplace
		ojaego.GetJeagoOverValueSum_Shop
	end if

%>
<script type='text/javascript'>

function reActMonthSummary() {

	var yyyymm = frm.yyyy1.value + "-" + frm.mm1.value;
	if (!confirm(yyyymm + ' 재고월령 새로고침 하시겠습니까?')){ return; }

	var popwin = window.open('do_stocksummary.asp?mode=stockovervalue&yyyymm=' + yyyymm,'reActMonthSummary','width=100,height=100');
	popwin.focus();
}

/*
function pop_exceldown() {
	<%
	'// 마지막 입고일(물류)
	'// db_summary.dbo.usp_Ten_monthly_Acc_SetLastIpgoDate_Logis
	'// 마지막 입고일(매장)
	'// db_summary.dbo.usp_Ten_monthly_Acc_SetLastIpgoDate_Shop
	'// 마지막 입고일(매입구분별)
	'// db_summary.[dbo].[sp_Ten_monthly_Maeip_Stockledger_Make]
	%>
	var popwin = window.open("/admin/newreport/monthlystock_overValue_csv.asp?exYYYY=<%'= yyyy1 %>&exMM=<%'= mm1 %>&stplace=<%'= stplace %>&sysorreal=<%'= sysorreal %>&bPriceGbn=<%'= bPriceGbn %>&mygubun=<%'= mygubun %>","pop_exceldown","width=600,height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}
 */

function pop_exceldown(){
	alert('다운로드중입니다. 기다려주세요.');
	document.frmexcel.target = "xLink";
	document.frmexcel.exYYYY.value = '<%= yyyy1 %>';
	document.frmexcel.exMM.value = '<%= mm1 %>';
	document.frmexcel.stplace.value = '<%= stplace %>';
	document.frmexcel.sysorreal.value = '<%= sysorreal %>';
	document.frmexcel.bPriceGbn.value = '<%= bPriceGbn %>';
	document.frmexcel.mygubun.value = '<%= mygubun %>';
	<% 'document.frmexcel.action = "/admin/newreport/monthlystock_overValue_csv.asp" %>
	document.frmexcel.action = "/admin/newreport/monthlystock_overValue_excel.asp"
	document.frmexcel.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
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
		&nbsp;&nbsp;
		<input type="checkbox" name="swSppPrc" value="Y" <%= CHKIIF(swSppPrc="Y","checked","") %> >공급가로 표시
	</td>

	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<% if (Not isViewUser) then %>
		<font color="#CC3333">재고구분:</font>
		<input type="radio" name="sysorreal" value="sys" <% if sysorreal="sys" then response.write "checked" %> >시스템재고
		<!--
		<input type="radio" name="sysorreal" value="real" <% if sysorreal="real" then response.write "checked" %> >실사재고
		-->
		&nbsp;&nbsp;
		<% end if %>

		<font color="#CC3333">매입구분:</font>
		<input type="radio" name="mwgubun" value="" <% if mwgubun="" then response.write "checked" %> >전체
		<input type="radio" name="mwgubun" value="M" <% if mwgubun="M" then response.write "checked" %> >매입(아이띵소(구) 출고위탁포함)
		<input type="radio" name="mwgubun" value="W" <% if mwgubun="W" then response.write "checked" %> >위탁
		<input type="radio" name="mwgubun" value="Z" <% if mwgubun="Z" then response.write "checked" %> >미지정
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">

	<font color="#CC3333">마이너스구분:</font>
	<input type="radio" name="minusinc" value="" <%= CHKIIF(minusinc="","checked","") %> >마이너스재고 포함(전체)
	<!--
	<input type="radio" name="minusinc" value="N" <%= CHKIIF(minusinc="N","checked","") %> >마이너스재고 제외
	-->
	&nbsp;&nbsp;
	<% if (Not isViewUser) then %>
	<font color="#CC3333">매입가기준:</font>
	<input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >작성시매입가
	<input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %>  >평균매입가
	&nbsp;
	<font color="#CC3333">산정기간:</font>
	<input type="radio" name="mygubun" value="M" <%= CHKIIF(mygubun="M","checked","") %>  >월별
	<input type="radio" name="mygubun" value="Y" <%= CHKIIF(mygubun="Y","checked","") %>  >연도별
	<% end if %>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<font color="#CC3333">재고위치:</font>
		<select name="stplace">
			<option value="L" <%= CHKIIF(stplace="L","selected" ,"") %> >물류</option>
			<option value="M" <%= CHKIIF(stplace="M","selected" ,"") %> >매장(매입구분별)</option>
			<option value="">---------</option>
			<option value="T" <%= CHKIIF(stplace="T","selected" ,"") %> >매장(물류입고일)</option>
			<option value="S" <%= CHKIIF(stplace="S","selected" ,"") %> >매장(매장입고일)</option>
		</select>
		&nbsp;
		<font color="#CC3333">부서구분:</font>
		<% Call drawSelectBoxBuseoGubunWith3PL("buseo", buseo) %>
		&nbsp;
		<font color="#CC3333">상품구분:</font>
		<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
		&nbsp;
		<% if (stplace = "S") or (stplace = "T") or (stplace = "M") then %>
			&nbsp;
			매장(매장재고 검색시) : <% Call drawSelectBoxAccShop(yyyy1 + "-" + mm1, "", "shopid", shopid) %>
			&nbsp;
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
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<br>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		* 최종입고월을 기준으로 <font color="red">재고월령</font>을 산정합니다.<br>
		* 재고월령이 1년을 넘는 상품의 경우 <font color="red">재고평가충당금(재고평가손실)</font>을 적용합니다.<br>
		* 재고월령이 1년-2년 사이인 경우 매입가 대비 50% 의 평가충당금을 산정합니다.<br>
		* 재고월령이 2년을 넘는 경우 매입가 대비 100% 의 평가충당금을 산정합니다.<br>
		* 매장(매입구분별) = 물류매입상품은 물류입고일, 그 이외 상품은 매장입고일.<br><br>
		* <font color="red">재고월령이 보이지 않는 경우</font><br>
		&nbsp; - 1. [통계]재고자산>>재고자산(월별) -- 재작성 (물류 / 매장)<br>
		&nbsp; - 2. [경영]재고자산>>재고자산(월별) FIX -- 삭제 후 복사<br>
	</td>
	<td align="right" valign="bottom">
		<% If stplace = "L" OR stplace = "T" OR stplace = "M" Then %>
			<input type="button" value="재고월령 다운로드" onclick="pop_exceldown();" class="button_s">
		<% End If %>
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="4">상품구분</td>
	<td rowspan="2" width="60">재고수량</td>

	<% if (mygubun = "Y") then %>
		<td rowspan="2" width="100"><%= yyyy1 %></td>
		<td rowspan="2" width="100"><%= (yyyy1 - 1) %></td>
		<td rowspan="2" width="100"><%= (yyyy1 - 2) %></td>
		<td rowspan="2" width="100">~ <%= (yyyy1 - 3) %></td>
	<% else %>
		<td rowspan="2" width="100">1개월~3개월</td>
		<td rowspan="2" width="100">4개월~6개월</td>
		<td rowspan="2" width="100">7개월~12개월</td>
		<td rowspan="2" width="100">13개월~18개월</td>
		<td rowspan="2" width="100">19개월~24개월</td>
		<td rowspan="2" width="100">2년초과</td>
	<% end if %>

	<td rowspan="2" width="100">NULL</td>
	<td rowspan="2" width="100">총계</td>
	<td rowspan="2" width="100">재고평가충당금</td>
	<td rowspan="2" width="100">재고액</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td >부서</td>
	<td >구분</td>
	<td >코드구분</td>
	<td >매입구분</td>
</tr>
<% for i=0 to ojaego.FResultCount-1 %>
<%
if (ojaego.FItemList(i).Fitemgubun <> "75") and (ojaego.FItemList(i).Fitemgubun <> "80") and (ojaego.FItemList(i).Fitemgubun <> "85") then
	totBuySum1 = totBuySum1 + ojaego.FItemList(i).FTotBuySum1
	totBuySum2 = totBuySum2 + ojaego.FItemList(i).FTotBuySum2
	totBuySum3 = totBuySum3 + ojaego.FItemList(i).FTotBuySum3
	totBuySum4 = totBuySum4 + ojaego.FItemList(i).FTotBuySum4
	totBuySum5 = totBuySum5 + ojaego.FItemList(i).FTotBuySum5
	totBuySum6 = totBuySum6 + ojaego.FItemList(i).FTotBuySum6
	totBuySum7 = totBuySum7 + ojaego.FItemList(i).FTotBuySum7
	totBuySum8 = totBuySum8 + ojaego.FItemList(i).FTotBuySum8
	totBuySum11 = totBuySum11 + ojaego.FItemList(i).FTotBuySum11
	totBuySum12 = totBuySum12 + ojaego.FItemList(i).FTotBuySum12
	totBuySum13 = totBuySum13 + ojaego.FItemList(i).FTotBuySum13
	totBuySum14 = totBuySum14 + ojaego.FItemList(i).FTotBuySum14
	totBuySum = totBuySum + ojaego.FItemList(i).FTotBuySum

	if (mygubun = "Y") then
		totOverValueSum = totOverValueSum + ojaego.FItemList(i).getOverValueStockPriceYear
	else
		totOverValueSum = totOverValueSum + ojaego.FItemList(i).getOverValueStockPrice
	end if

	tot_totStockNo = tot_totStockNo + ojaego.FItemList(i).FtotStockNo

	subTotBuySum1 = subTotBuySum1 + ojaego.FItemList(i).FTotBuySum1
	subTotBuySum2 = subTotBuySum2 + ojaego.FItemList(i).FTotBuySum2
	subTotBuySum3 = subTotBuySum3 + ojaego.FItemList(i).FTotBuySum3
	subTotBuySum4 = subTotBuySum4 + ojaego.FItemList(i).FTotBuySum4
	subTotBuySum5 = subTotBuySum5 + ojaego.FItemList(i).FTotBuySum5
	subTotBuySum6 = subTotBuySum6 + ojaego.FItemList(i).FTotBuySum6
	subTotBuySum7 = subTotBuySum7 + ojaego.FItemList(i).FTotBuySum7
	subTotBuySum8 = subTotBuySum8 + ojaego.FItemList(i).FTotBuySum8
	subTotBuySum11 = subTotBuySum11 + ojaego.FItemList(i).FTotBuySum11
	subTotBuySum12 = subTotBuySum12 + ojaego.FItemList(i).FTotBuySum12
	subTotBuySum13 = subTotBuySum13 + ojaego.FItemList(i).FTotBuySum13
	subTotBuySum14 = subTotBuySum14 + ojaego.FItemList(i).FTotBuySum14
	subTotBuySum = subTotBuySum + ojaego.FItemList(i).FTotBuySum

	if (mygubun = "Y") then
		subTotOverValueSum = subTotOverValueSum + ojaego.FItemList(i).getOverValueStockPriceYear
	else
		subTotOverValueSum = subTotOverValueSum + ojaego.FItemList(i).getOverValueStockPrice
	end if

	sub_totStockNo = sub_totStockNo + ojaego.FItemList(i).FtotStockNo

	iURL = "monthlystock_overValue_detail_2.asp?menupos="& menupos &"&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
	iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype &"&stplace="&stplace &"&shopid="&shopid&"&swSppPrc="&swSppPrc
	iURL = iURL + "&sysorreal=" & sysorreal + "&mygubun=" & mygubun & "&etcjungsantype="&etcjungsantype
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getBusiName %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getITemGubunName %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).Fitemgubun %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
		<td align="right"><a href="<%= iURL %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FtotStockNo,0) %></a></td>

		<% if (mygubun = "Y") then %>
		<td align="right"><a href="<%= iURL + "&monthGubun=11" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum11,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=12" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum12,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=13" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum13,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=14" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum14,0) %></a></td>
		<% else %>
		<td align="right"><a href="<%= iURL + "&monthGubun=1" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum1,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=2" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum2,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=3" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum3,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=7" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum7,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=8" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum8,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=5" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum5,0) %></a></td>
		<% end if %>

		<td align="right"><a href="<%= iURL + "&monthGubun=6" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum6,0) %></a></td>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>

		<% if (mygubun = "Y") then %>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).getOverValueStockPriceYear,0) %></td>
		<td align="right"><%= FormatNumber((ojaego.FItemList(i).FTotBuySum - ojaego.FItemList(i).getOverValueStockPriceYear),0) %></td>
		<% else %>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).getOverValueStockPrice,0) %></td>
		<td align="right"><%= FormatNumber((ojaego.FItemList(i).FTotBuySum - ojaego.FItemList(i).getOverValueStockPrice),0) %></td>
		<% end if %>
	</tr>
<% end if %>
<% next %>
<tr align="center" bgcolor="#EEFFEE">
	<td></td>
	<td>상품소계</td>
	<td></td>
	<td></td>
	<td align="right"><%= FormatNumber(sub_totStockNo,0) %></td>

	<% if (mygubun = "Y") then %>
		<td align="right" ><%= FormatNumber(subTotBuySum11,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum12,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum13,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum14,0) %></td>
	<% else %>
		<td align="right" ><%= FormatNumber(subTotBuySum1,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum2,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum3,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum7,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum8,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum5,0) %></td>
	<% end if %>

	<td align="right" ><%= FormatNumber(subTotBuySum6,0) %></td>
	<td align="right" ><b><%= FormatNumber(subTotBuySum,0) %></b></td>
	<td align="right" ><%= FormatNumber(subTotOverValueSum,0) %></td>
	<td align="right" ><b><%= FormatNumber(subTotBuySum - subTotOverValueSum,0) %></b></td>
</tr>
<tr  bgcolor="#FFFFFF">
	<td colspan="15"></td>
</tr>
<%
subTotBuySum1 = 0
subTotBuySum2 = 0
subTotBuySum3 = 0
subTotBuySum4 = 0
subTotBuySum5 = 0
subTotBuySum6 = 0
subTotBuySum7 = 0
subTotBuySum8 = 0
subTotBuySum11 = 0
subTotBuySum12 = 0
subTotBuySum13 = 0
subTotBuySum14 = 0
subTotBuySum = 0
subTotOverValueSum = 0
sub_totStockNo = 0
%>
<% for i=0 to ojaego.FResultCount-1 %>
<%
if (ojaego.FItemList(i).Fitemgubun = "75") or (ojaego.FItemList(i).Fitemgubun = "80") or (ojaego.FItemList(i).Fitemgubun = "85") then
	totBuySum1 = totBuySum1 + ojaego.FItemList(i).FTotBuySum1
	totBuySum2 = totBuySum2 + ojaego.FItemList(i).FTotBuySum2
	totBuySum3 = totBuySum3 + ojaego.FItemList(i).FTotBuySum3
	totBuySum4 = totBuySum4 + ojaego.FItemList(i).FTotBuySum4
	totBuySum5 = totBuySum5 + ojaego.FItemList(i).FTotBuySum5
	totBuySum6 = totBuySum6 + ojaego.FItemList(i).FTotBuySum6
	totBuySum7 = totBuySum7 + ojaego.FItemList(i).FTotBuySum7
	totBuySum8 = totBuySum8 + ojaego.FItemList(i).FTotBuySum8
	totBuySum11 = totBuySum11 + ojaego.FItemList(i).FTotBuySum11
	totBuySum12 = totBuySum12 + ojaego.FItemList(i).FTotBuySum12
	totBuySum13 = totBuySum13 + ojaego.FItemList(i).FTotBuySum13
	totBuySum14 = totBuySum14 + ojaego.FItemList(i).FTotBuySum14
	totBuySum = totBuySum + ojaego.FItemList(i).FTotBuySum

	if (mygubun = "Y") then
		totOverValueSum = totOverValueSum + ojaego.FItemList(i).getOverValueStockPriceYear
	else
		totOverValueSum = totOverValueSum + ojaego.FItemList(i).getOverValueStockPrice
	end if

	tot_totStockNo = tot_totStockNo + ojaego.FItemList(i).FtotStockNo

	subTotBuySum1 = subTotBuySum1 + ojaego.FItemList(i).FTotBuySum1
	subTotBuySum2 = subTotBuySum2 + ojaego.FItemList(i).FTotBuySum2
	subTotBuySum3 = subTotBuySum3 + ojaego.FItemList(i).FTotBuySum3
	subTotBuySum4 = subTotBuySum4 + ojaego.FItemList(i).FTotBuySum4
	subTotBuySum5 = subTotBuySum5 + ojaego.FItemList(i).FTotBuySum5
	subTotBuySum6 = subTotBuySum6 + ojaego.FItemList(i).FTotBuySum6
	subTotBuySum7 = subTotBuySum7 + ojaego.FItemList(i).FTotBuySum7
	subTotBuySum8 = subTotBuySum8 + ojaego.FItemList(i).FTotBuySum8
	subTotBuySum11 = subTotBuySum11 + ojaego.FItemList(i).FTotBuySum11
	subTotBuySum12 = subTotBuySum12 + ojaego.FItemList(i).FTotBuySum12
	subTotBuySum13 = subTotBuySum13 + ojaego.FItemList(i).FTotBuySum13
	subTotBuySum14 = subTotBuySum14 + ojaego.FItemList(i).FTotBuySum14
	subTotBuySum = subTotBuySum + ojaego.FItemList(i).FTotBuySum

	if (mygubun = "Y") then
		subTotOverValueSum = subTotOverValueSum + ojaego.FItemList(i).getOverValueStockPriceYear
	else
		subTotOverValueSum = subTotOverValueSum + ojaego.FItemList(i).getOverValueStockPrice
	end if

	sub_totStockNo = sub_totStockNo + ojaego.FItemList(i).FtotStockNo

	iURL = "monthlystock_overValue_detail_2.asp?menupos="& menupos &"&mwgubun="& ojaego.FItemList(i).FMaeIpGubun &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing="& isusing &"&newitem="& newitem &"&itemgubun="&ojaego.FItemList(i).Fitemgubun&"&vatyn="&vatyn
	iURL = iURL + "&minusinc="&minusinc&"&bPriceGbn="&bPriceGbn&"&buseo="&ojaego.FItemList(i).FtargetGbn&"&purchasetype="&purchasetype &"&stplace="&stplace &"&shopid="&shopid&"&swSppPrc="&swSppPrc
	iURL = iURL + "&sysorreal=" & sysorreal + "&mygubun=" & mygubun & "&etcjungsantype="&etcjungsantype
%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getBusiName %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getITemGubunName %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).Fitemgubun %></a></td>
		<td><a href="<%= iURL %>" target="_blank"><%= ojaego.FItemList(i).getMaeipGubunName %></a></td>
		<td align="right"><a href="<%= iURL %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FtotStockNo,0) %></a></td>

		<% if (mygubun = "Y") then %>
		<td align="right"><a href="<%= iURL + "&monthGubun=11" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum11,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=12" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum12,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=13" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum13,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=14" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum14,0) %></a></td>
		<% else %>
		<td align="right"><a href="<%= iURL + "&monthGubun=1" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum1,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=2" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum2,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=3" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum3,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=7" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum7,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=8" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum8,0) %></a></td>
		<td align="right"><a href="<%= iURL + "&monthGubun=5" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum5,0) %></a></td>
		<% end if %>

		<td align="right"><a href="<%= iURL + "&monthGubun=6" %>" target="_blank"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum6,0) %></a></td>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).FTotBuySum,0) %></td>

		<% if (mygubun = "Y") then %>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).getOverValueStockPriceYear,0) %></td>
		<td align="right"><%= FormatNumber((ojaego.FItemList(i).FTotBuySum - ojaego.FItemList(i).getOverValueStockPriceYear),0) %></td>
		<% else %>
		<td align="right"><%= FormatNumber(ojaego.FItemList(i).getOverValueStockPrice,0) %></td>
		<td align="right"><%= FormatNumber((ojaego.FItemList(i).FTotBuySum - ojaego.FItemList(i).getOverValueStockPrice),0) %></td>
		<% end if %>
	</tr>
<% end if %>
<% next %>
<tr align="center" bgcolor="#EEFFEE">
	<td></td>
	<td>저장품소계</td>
	<td></td>
	<td></td>
	<td align="right"><%= FormatNumber(sub_totStockNo,0) %></td>

	<% if (mygubun = "Y") then %>
		<td align="right" ><%= FormatNumber(subTotBuySum11,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum12,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum13,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum14,0) %></td>
	<% else %>
		<td align="right" ><%= FormatNumber(subTotBuySum1,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum2,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum3,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum7,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum8,0) %></td>
		<td align="right" ><%= FormatNumber(subTotBuySum5,0) %></td>
	<% end if %>

	<td align="right" ><%= FormatNumber(subTotBuySum6,0) %></td>
	<td align="right" ><b><%= FormatNumber(subTotBuySum,0) %></b></td>
	<td align="right" ><%= FormatNumber(subTotOverValueSum,0) %></td>
	<td align="right" ><b><%= FormatNumber(subTotBuySum - subTotOverValueSum,0) %></b></td>
</tr>
<tr  bgcolor="#FFFFFF">
	<td colspan="15"></td>
</tr>
<%
subTotBuySum1 = 0
subTotBuySum2 = 0
subTotBuySum3 = 0
subTotBuySum4 = 0
subTotBuySum5 = 0
subTotBuySum6 = 0
subTotBuySum7 = 0
subTotBuySum8 = 0
subTotBuySum11 = 0
subTotBuySum12 = 0
subTotBuySum13 = 0
subTotBuySum14 = 0
subTotBuySum = 0
subTotOverValueSum = 0
sub_totStockNo = 0
%>
<tr align="center" bgcolor="#FFFFFF">
	<td></td>
	<td>총계</td>
	<td></td>
	<td></td>
	<td align="right"><%= FormatNumber(tot_totStockNo,0) %></td>

	<% if (mygubun = "Y") then %>
		<td align="right" ><%= FormatNumber(totBuySum11,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum12,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum13,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum14,0) %></td>
	<% else %>
		<td align="right" ><%= FormatNumber(totBuySum1,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum2,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum3,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum7,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum8,0) %></td>
		<td align="right" ><%= FormatNumber(totBuySum5,0) %></td>
	<% end if %>

	<td align="right" ><%= FormatNumber(totBuySum6,0) %></td>
	<td align="right" ><b><%= FormatNumber(totBuySum,0) %></b></td>
	<td align="right" ><%= FormatNumber(totOverValueSum,0) %></td>
	<td align="right" ><b><%= FormatNumber(totBuySum - totOverValueSum,0) %></b></td>
</tr>
</table>

<form name="frmexcel" method="post" style="margin:0px;">
<input type="hidden" name="exYYYY">
<input type="hidden" name="exMM">
<input type="hidden" name="stplace">
<input type="hidden" name="sysorreal">
<input type="hidden" name="bPriceGbn">
<input type="hidden" name="mygubun">
</form>
<% IF application("Svr_Info")="Dev" THEN %>
	<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="300"></iframe>
<% else %>
	<iframe name="xLink" id="xLink" frameborder="0" width="0" height="0"></iframe>
<% end if %>

<%
set ojaego = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
