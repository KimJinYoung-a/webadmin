<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 재고자산(월별)
' History : 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls.asp"-->
<%



Dim CCADMIN : CCADMIN = C_ADMIN_AUTH

dim research, i
dim yyyy1,mm1, yyyymm1, makerid, showsuply, meaipTp, showShopid
dim stockPlace, shopid
dim targetGbn, itemgubun
dim bPriceGbn
dim showUpbae

research    = requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1       	= requestCheckvar(request("mm1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsuply   = requestCheckvar(request("showsuply"),10)
showShopid  = requestCheckvar(request("showShopid"),10)
shopid    	= requestCheckvar(request("shopid"),32)
meaipTp     = requestCheckvar(request("meaipTp"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
targetGbn   = requestCheckvar(request("targetGbn"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
showUpbae   = requestCheckvar(request("showUpbae"),10)

''if (stockPlace="") then stockPlace="L"

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

if (research="") and (bPriceGbn = "") then
    bPriceGbn = "V"
	showsuply = "on"
end if

yyyymm1 = yyyy1 + "-" + mm1

'if (yyyymm1="2016-05") and (session("ssBctID")<>"icommang") then
'    response.write "수정중"
'    dbget.close() : response.end
'end if

dim oCMonthlyMaeipLedge
set oCMonthlyMaeipLedge = new CMonthlyMaeipLedge

oCMonthlyMaeipLedge.FRectYYYYMM = yyyymm1
oCMonthlyMaeipLedge.FRectStockPlace = stockPlace
oCMonthlyMaeipLedge.FRectMakerid = makerid
oCMonthlyMaeipLedge.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
oCMonthlyMaeipLedge.FRectMeaipTp = meaipTp
oCMonthlyMaeipLedge.FRectShopid = shopid
oCMonthlyMaeipLedge.FRectItemgubun = itemgubun
oCMonthlyMaeipLedge.FRectTargetGbn = targetGbn
oCMonthlyMaeipLedge.FRectShowShopid = showShopid
oCMonthlyMaeipLedge.FRectPriceGubun = bPriceGbn
oCMonthlyMaeipLedge.FRectShowUpbae = showUpbae

oCMonthlyMaeipLedge.GetMaeipLedgeSUM

dim oCMonthlyMaeipJungsan
set oCMonthlyMaeipJungsan = new CMonthlyMaeipLedge

oCMonthlyMaeipJungsan.FRectYYYYMM = yyyymm1
oCMonthlyMaeipJungsan.FRectStockPlace = stockPlace
oCMonthlyMaeipJungsan.FRectMakerid = makerid
oCMonthlyMaeipJungsan.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
oCMonthlyMaeipJungsan.FRectMeaipTp = meaipTp
oCMonthlyMaeipJungsan.FRectShopid = shopid
oCMonthlyMaeipJungsan.FRectItemgubun = itemgubun
oCMonthlyMaeipJungsan.FRectTargetGbn = targetGbn
oCMonthlyMaeipJungsan.FRectShowShopid = showShopid

oCMonthlyMaeipJungsan.GetMaeipJungsanSum


dim oCMonthlyIpMaeipJungsan
set oCMonthlyIpMaeipJungsan = new CMonthlyMaeipLedge

oCMonthlyIpMaeipJungsan.FRectYYYYMM = yyyymm1
oCMonthlyIpMaeipJungsan.FRectStockPlace = stockPlace
oCMonthlyIpMaeipJungsan.FRectMakerid = makerid
oCMonthlyIpMaeipJungsan.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
oCMonthlyIpMaeipJungsan.FRectMeaipTp = meaipTp
oCMonthlyIpMaeipJungsan.FRectShopid = shopid
oCMonthlyIpMaeipJungsan.FRectItemgubun = itemgubun
oCMonthlyIpMaeipJungsan.FRectTargetGbn = targetGbn
oCMonthlyIpMaeipJungsan.FRectShowShopid = showShopid
oCMonthlyIpMaeipJungsan.FRectOnlyIpgoMeaip = "on"

if (CCADMIN) then
    oCMonthlyIpMaeipJungsan.GetMaeipJungsanSum
end if

dim totprevSysStockNo, totprevSysStockSum, totIpgoNo, totIpgoSum, totSellNo, totSellSum, totOffChulNo, totOffChulSum, totEtcChulNo, totEtcChulSum
dim totCsNo, totCsSum, totLossChulNo, totLossChulSum, totcurSysStockNo, totcurSysStockSum, totcurErrRealCheckNo, totcurErrRealCheckSum
dim diff, totdiff
dim diffSum, totdiffSum
dim totMoveNo, totMoveSum
dim totErrNo, totErrSum

dim totprevSysStockNo2, totprevSysStockSum2, totIpgoNo2, totIpgoSum2, totSellNo2, totSellSum2, totOffChulNo2, totOffChulSum2, totEtcChulNo2, totEtcChulSum2
dim totCsNo2, totCsSum2, totLossChulNo2, totLossChulSum2, totcurSysStockNo2, totcurSysStockSum2
dim totMoveNo2, totMoveSum2
dim totErrNo2, totErrSum2

dim jtotprevSysStockNo, jtotprevSysStockSum, jtotIpgoNo, jtotIpgoSum, jtotSellNo, jtotSellSum, jtotOffChulNo, jtotOffChulSum, jtotEtcChulNo, jtotEtcChulSum
dim jtotCsNo, jtotCsSum, jtotLossChulNo, jtotLossChulSum, jtotcurSysStockNo, jtotcurSysStockSum, jtotcurErrRealCheckNo, jtotcurErrRealCheckSum
dim jtotMoveNo, jtotMoveSum

dim jIptotprevSysStockNo, jIptotprevSysStockSum, jIptotIpgoNo, jIptotIpgoSum, jIptotSellNo, jIptotSellSum, jIptotOffChulNo, jIptotOffChulSum, jIptotEtcChulNo, jIptotEtcChulSum
dim jIptotCsNo, jIptotCsSum, jIptotLossChulNo, jIptotLossChulSum, jIptotcurSysStockNo, jIptotcurSysStockSum, jIptotcurErrRealCheckNo, jIptotcurErrRealCheckSum
dim jIptotMoveNo, jIptotMoveSum

dim iURL, iURLEtc


dim stockPlaceName, PriceGbnName
select case stockPlace
	case "L"
		stockPlaceName = "물류"
	case "S"
		stockPlaceName = "매장"
	case else
		''
end select

select case bPriceGbn
	case "V"
		PriceGbnName = "평균매입가"
	case "P"
		PriceGbnName = "작성시매입가"
	case else
		''
end select

dim itemgubunColNUM : itemgubunColNUM = 4
if (showShopid <> "") then
	itemgubunColNUM = 5
end if

%>
<script language='javascript'>

function jsstockDown(){
	document.frm.action="/admin/newreport/monthlystock_excel.asp";
	frm.target = "exceldown";
	document.frm.submit();
	frm.target = "";
	frm.action = ""
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" action="" target="" style="margin:0px;" >
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			<font color="#CC3333">년/월 :</font> <% DrawYMBox yyyy1,mm1 %>
			&nbsp;&nbsp;
			<font color="#CC3333">브랜드:</font> <%	drawSelectBoxDesignerWithName "makerid", makerid %>
			매장 : <% drawSelectBoxOffShopNotUsingAll "shopid",shopid %>
			&nbsp;&nbsp;
			<input type="checkbox" name="showsuply" value="on" <%= CHKIIF(showsuply="on","checked","") %> >공급가로 표시
			&nbsp;&nbsp;
			<input type="checkbox" name="showShopid" value="on" <%= CHKIIF(showShopid="on","checked","") %> >매장 표시
			&nbsp;&nbsp;
			<input type="checkbox" name="showUpbae" value="on" <%= CHKIIF(showUpbae="on","checked","") %> >업배상품만 표시
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.target='';document.frm.action='';document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    <font color="#CC3333">재고위치:</font>
		    <select name="stockPlace">
		        <option value="" <%= CHKIIF(stockPlace="","selected" ,"") %> >전체</option>
        		<option value="L" <%= CHKIIF(stockPlace="L","selected" ,"") %> >물류</option>
        		<option value="S" <%= CHKIIF(stockPlace="S","selected" ,"") %> >매장</option>
				<option value="T" <%= CHKIIF(stockPlace="T","selected" ,"") %> >띵소</option>
				<option value=""">---------</option>
				<option value="O" <%= CHKIIF(stockPlace="O","selected" ,"") %> >온라인 매입정산</option>
				<option value="N" <%= CHKIIF(stockPlace="N","selected" ,"") %> >온라인 매입정산(공제불가)</option>
				<option value="F" <%= CHKIIF(stockPlace="F","selected" ,"") %> >오프 매입정산</option>
				<option value="A" <%= CHKIIF(stockPlace="A","selected" ,"") %> >핑거스 매입정산</option>
				<option value=""">---------</option>
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
			<% drawSelectBoxItemGubun "itemgubun", itemgubun %>
			&nbsp;&nbsp;
			<font color="#CC3333">매입가기준:</font>
			<input type="radio" name="bPriceGbn" value="P" <%= CHKIIF(bPriceGbn="P","checked","") %>  >작성시매입가
			<input type="radio" name="bPriceGbn" value="V" <%= CHKIIF(bPriceGbn="V","checked","") %>  >평균매입가
	    </td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<% if True or CCADMIN then %>
<script  language='javascript'>
function remakeSum(){
    var actFrm = document.frmAct;

    if (actFrm.atype.value.length<1){
        alert('작성구분을 선택하세요.');
        actFrm.ptype.focus();
        return;
    }

    if (actFrm.ptype.value.length<1){
        alert('재고위치를 선택하세요.');
        actFrm.ptype.focus();
        return;
    }

    if (confirm('<%=yyyy1%>-<%=mm1%> 작성구분['+ actFrm.atype.value +'], 재고위치['+ actFrm.ptype.value +'] 생성하시겠습니까?')){
        var popwin = window.open("","remakeSum","width=600,height=600");

	    actFrm.target="remakeSum";
	    actFrm.submit();
	    popwin.focus();
    }
}

</script>
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="#FFFFFF">
<form name="frmAct" method="post" action="<%=stsAdmURL%>/admin/newreport/do_stocksummary.asp">
<input type="hidden" name="mode" value="meaipsummake">
<input type="hidden" name="yyyymm" value="<%=yyyy1%>-<%=mm1%>">
<tr>
    <td>
    <select name="atype">
    <option value="">작성구분선택
    <option value="S">매입재고
    <option value="J">정산
    </select>

    <select name="ptype">
		<option value="">재고위치선택</option>
		<option value="A">전체</option>
		<option value="L">물류</option>
		<option value="S">매장</option>
		<option value="J">위탁상품 매입정산</option>
    </select>

    <input type="button" value="생성" onClick="remakeSum()">
	</td>
	<td align="right">
		<% if (CCADMIN) then %>
			*엑셀받기 기능은 <b>재고자산(월별) FIX</b> 메뉴에서만 가능합니다.
		<% end if %>
    </td>
</tr>
</form>
</table>
<p>
<% end if %>
<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="5" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td colspan="<%= itemgubunColNUM %>">상품구분</td>
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
		<td >구분</td>
	    <td >코드<br>구분</td>
	    <td >매입구분</td>
	    <td >재고<br>위치</td>
		<% if (showShopid <> "") then %>
		<td >샵아이디</td>
		<% end if %>
    	<td width="50">수량</td>
    	<td width="80">금액</td>
    	<td width="50">수량</td>
    	<td width="70">금액</td>
    	<td width="50">수량</td>
    	<td width="70">금액</td>
    	<td width="50">수량</td>
    	<td width="70">금액</td>
    	<td width="50">수량</td>
    	<td width="70">금액</td>
    	<td width="50">수량</td>
    	<td width="70">금액</td>
    	<td width="50">수량</td>
    	<td width="70">금액</td>
    	<td width="50">수량</td>
    	<td width="70">금액</td>
    	<td width="50">수량</td>
    	<td width="70">금액</td>
    	<td width="50">수량</td>
    	<td width="80">금액</td>
    </tr>
    <% for i=0 to oCMonthlyMaeipLedge.FResultCount-1 %>
	<% if (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun <> "75") and (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun <> "80") and (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun <> "85") then %>
    <%

	totprevSysStockNo       	= totprevSysStockNo + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo
	totprevSysStockSum       	= totprevSysStockSum + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum

	totIpgoNo       			= totIpgoNo + oCMonthlyMaeipLedge.FItemList(i).getIpgoNo
	totIpgoSum       			= totIpgoSum + oCMonthlyMaeipLedge.FItemList(i).getIpgoSum

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
    'diffSum = oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum + oCMonthlyMaeipLedge.FItemList(i).getIpgoSum + oCMonthlyMaeipLedge.FItemList(i).getMoveSum + oCMonthlyMaeipLedge.FItemList(i).FSellSum + oCMonthlyMaeipLedge.FItemList(i).FOffChulSum + oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum + oCMonthlyMaeipLedge.FItemList(i).FCsSum + oCMonthlyMaeipLedge.FItemList(i).FLossChulSum - oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum
	diff = oCMonthlyMaeipLedge.FItemList(i).getDiffNo
	diffSum = oCMonthlyMaeipLedge.FItemList(i).getDiffSum
	totdiff = totdiff + diff
	totdiffSum = totdiffSum + diffSum

	totErrNo = totErrNo + oCMonthlyMaeipLedge.FItemList(i).getTotErrNo
    totErrSum = totErrSum + oCMonthlyMaeipLedge.FItemList(i).getTotErrSum


	iURL = "monthlyMaeipLedge_summaryDetail.asp?menupos="& menupos
	iURL = iURL + "&yyyy1="& yyyy1 &"&mm1="& mm1 &"&makerid="& makerid &"&showsuply="&showsuply&"&meaipTp="&oCMonthlyMaeipLedge.FItemList(i).Flastmwdiv
    iURL = iURL + "&itemgubun="&oCMonthlyMaeipLedge.FItemList(i).Fitemgubun&"&targetGbn="&oCMonthlyMaeipLedge.FItemList(i).FtargetGbn
    iURL = iURL + "&stockPlace="&oCMonthlyMaeipLedge.FItemList(i).FstockPlace&"&shopid="&oCMonthlyMaeipLedge.FItemList(i).Fshopid&"&stype=S"
    iURL = iURL + "&bPriceGbn="&bPriceGbn

	iURLEtc = "monthlystock_etcChulgoList.asp?menupos="& menupos &"&dtype=mk" & "&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing=&newitem=&itemgubun="& oCMonthlyMaeipLedge.FItemList(i).Fitemgubun &"&vatyn="
	if (oCMonthlyMaeipLedge.FItemList(i).FstockPlace <> "L") and (oCMonthlyMaeipLedge.FItemList(i).FstockPlace <> "S") then
		iURLEtc = iURLEtc + "&minusinc=&bPriceGbn="&bPriceGbn&"&buseo="& oCMonthlyMaeipLedge.FItemList(i).FtargetGbn &"&purchasetype=&mwgubun=W&stplace=L&shopid="& oCMonthlyMaeipLedge.FItemList(i).Fshopid &"&etcjungsantype="
	else
		iURLEtc = iURLEtc + "&minusinc=&bPriceGbn="&bPriceGbn&"&buseo="& oCMonthlyMaeipLedge.FItemList(i).FtargetGbn &"&purchasetype=&mwgubun=" & oCMonthlyMaeipLedge.FItemList(i).Flastmwdiv & "&stplace="& oCMonthlyMaeipLedge.FItemList(i).FstockPlace &"&shopid="& oCMonthlyMaeipLedge.FItemList(i).Fshopid &"&etcjungsantype="
	end if
    iURLEtc=iURLEtc&"&sysorreal=sys"
    %>
    <tr align="right" bgcolor="#FFFFFF" >
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= GetItemGubunName(oCMonthlyMaeipLedge.FItemList(i).Fitemgubun) %></a></td>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %></a></td>
        <td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipLedge.FItemList(i).getMeaipTypeName %></a></td>
        <td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipLedge.FItemList(i).FstockPlace%></a></td>
		<% if (showShopid <> "") then %>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipLedge.FItemList(i).Fshopid %></a></td>
		<% end if %>
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

		<td><a href="<%= iURLEtc %>&chulgogubun=etc" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo,0) %></a></td>
		<td><a href="<%= iURLEtc %>&chulgogubun=etc" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum,0) %></a></td>

		<% if (oCMonthlyMaeipLedge.FItemList(i).FstockPlace = "O") or (oCMonthlyMaeipLedge.FItemList(i).FstockPlace = "N") then %>
		<td><a href="<%= iURLEtc %>&chulgogubun=<%= CHKIIF(oCMonthlyMaeipLedge.FItemList(i).FstockPlace="N", "etc2", "etc3") %>" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulNo,0) %></a></td>
		<td><a href="<%= iURLEtc %>&chulgogubun=<%= CHKIIF(oCMonthlyMaeipLedge.FItemList(i).FstockPlace="N", "etc2", "etc3") %>" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulSum,0) %></a></td>
		<% else %>
		<td><a href="<%= iURLEtc %>&chulgogubun=" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulNo,0) %></a></td>
		<td><a href="<%= iURLEtc %>&chulgogubun=" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulSum,0) %></a></td>
		<% end if %>

		<td>
			<% if (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun = "10") and (oCMonthlyMaeipLedge.FItemList(i).FstockPlace = "L") then %>
			<a href="/cscenter/action/cschulgolist.asp?menupos=1768&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>" target="_blank">
			<% end if %>
			<%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FCsNo,0) %>
		</td>
		<td>
			<% if (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun = "10") and (oCMonthlyMaeipLedge.FItemList(i).FstockPlace = "L") then %>
			<a href="/cscenter/action/cschulgolist.asp?menupos=1768&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>" target="_blank">
			<% end if %>
			<%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FCsSum,0) %>
		</td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getTotErrNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getTotErrSum,0) %></td>

		<td><b><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo,0) %></b></td>
		<td><b><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum,0) %></b></td>
    </tr>
	<% end if %>
	<% next %>

    <tr align="center" bgcolor="#EEEEEE">
    	<td colspan="<%= itemgubunColNUM %>">상품소계</td>
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

		<td align="right" ><b><%= FormatNumber(totcurSysStockNo,0) %></b></td>
		<td align="right" ><b><%= FormatNumber(totcurSysStockSum,0) %></b></td>
    </tr>

    <% for i=0 to oCMonthlyMaeipLedge.FResultCount-1 %>
	<% if (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun = "75") or (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun = "80") or (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun = "85") then %>
    <%

	totprevSysStockNo       	= totprevSysStockNo + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo
	totprevSysStockSum       	= totprevSysStockSum + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum

	totIpgoNo       			= totIpgoNo + oCMonthlyMaeipLedge.FItemList(i).getIpgoNo
	totIpgoSum       			= totIpgoSum + oCMonthlyMaeipLedge.FItemList(i).getIpgoSum

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
    'diffSum = oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum + oCMonthlyMaeipLedge.FItemList(i).getIpgoSum + oCMonthlyMaeipLedge.FItemList(i).getMoveSum + oCMonthlyMaeipLedge.FItemList(i).FSellSum + oCMonthlyMaeipLedge.FItemList(i).FOffChulSum + oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum + oCMonthlyMaeipLedge.FItemList(i).FCsSum + oCMonthlyMaeipLedge.FItemList(i).FLossChulSum - oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum
	diff = oCMonthlyMaeipLedge.FItemList(i).getDiffNo
	diffSum = oCMonthlyMaeipLedge.FItemList(i).getDiffSum
	totdiff = totdiff + diff
	totdiffSum = totdiffSum + diffSum

	totErrNo = totErrNo + oCMonthlyMaeipLedge.FItemList(i).getTotErrNo
    totErrSum = totErrSum + oCMonthlyMaeipLedge.FItemList(i).getTotErrSum

	totprevSysStockNo2       	= totprevSysStockNo2 + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo
	totprevSysStockSum2       	= totprevSysStockSum2 + oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum

	totIpgoNo2       			= totIpgoNo2 + oCMonthlyMaeipLedge.FItemList(i).getIpgoNo
	totIpgoSum2       			= totIpgoSum2 + oCMonthlyMaeipLedge.FItemList(i).getIpgoSum

    totMoveNo2       			= totMoveNo2 + oCMonthlyMaeipLedge.FItemList(i).getMoveNo
    totMoveSum2       			= totMoveSum2 + oCMonthlyMaeipLedge.FItemList(i).getMoveSum

	totSellNo2       			= totSellNo2 + oCMonthlyMaeipLedge.FItemList(i).FSellNo
	totSellSum2       			= totSellSum2 + oCMonthlyMaeipLedge.FItemList(i).FSellSum

	totOffChulNo2       		= totOffChulNo2 + oCMonthlyMaeipLedge.FItemList(i).FOffChulNo
	totOffChulSum2       		= totOffChulSum2 + oCMonthlyMaeipLedge.FItemList(i).FOffChulSum

	totEtcChulNo2       		= totEtcChulNo2 + oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo
	totEtcChulSum2       		= totEtcChulSum2 + oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum

	totLossChulNo2       		= totLossChulNo2 + oCMonthlyMaeipLedge.FItemList(i).FLossChulNo
	totLossChulSum2       		= totLossChulSum2 + oCMonthlyMaeipLedge.FItemList(i).FLossChulSum

	totCsNo2       				= totCsNo2 + oCMonthlyMaeipLedge.FItemList(i).FCsNo
	totCsSum2       			= totCsSum2 + oCMonthlyMaeipLedge.FItemList(i).FCsSum

	totcurSysStockNo2       	= totcurSysStockNo2 + oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo
	totcurSysStockSum2       	= totcurSysStockSum2 + oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum

	totErrNo2 = totErrNo2 + oCMonthlyMaeipLedge.FItemList(i).getTotErrNo
    totErrSum2 = totErrSum2 + oCMonthlyMaeipLedge.FItemList(i).getTotErrSum

	iURL = "monthlyMaeipLedge_summaryDetail.asp?menupos="& menupos
	iURL = iURL + "&yyyy1="& yyyy1 &"&mm1="& mm1 &"&makerid="& makerid &"&showsuply="&showsuply&"&meaipTp="&oCMonthlyMaeipLedge.FItemList(i).Flastmwdiv
    iURL = iURL + "&itemgubun="&oCMonthlyMaeipLedge.FItemList(i).Fitemgubun&"&targetGbn="&oCMonthlyMaeipLedge.FItemList(i).FtargetGbn
    iURL = iURL + "&stockPlace="&oCMonthlyMaeipLedge.FItemList(i).FstockPlace&"&shopid="&oCMonthlyMaeipLedge.FItemList(i).Fshopid&"&stype=S"
    iURL = iURL + "&bPriceGbn="&bPriceGbn

	iURLEtc = "monthlystock_etcChulgoList.asp?menupos="& menupos &"&dtype=mk&mwgubun="& oCMonthlyMaeipLedge.FItemList(i).Flastmwdiv &"&yyyy1="& yyyy1 &"&mm1="& mm1 &"&isusing=&newitem=&itemgubun="& oCMonthlyMaeipLedge.FItemList(i).Fitemgubun &"&vatyn="
    iURLEtc = iURLEtc + "&minusinc=&bPriceGbn="&bPriceGbn&"&buseo="& oCMonthlyMaeipLedge.FItemList(i).FtargetGbn &"&purchasetype=&stplace="& oCMonthlyMaeipLedge.FItemList(i).FstockPlace &"&shopid="& oCMonthlyMaeipLedge.FItemList(i).Fshopid &"&etcjungsantype="
    iURLEtc=iURLEtc&"&sysorreal=sys"
    %>
    <tr align="right" bgcolor="#FFFFFF" >
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= GetItemGubunName(oCMonthlyMaeipLedge.FItemList(i).Fitemgubun) %></a></td>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipLedge.FItemList(i).Fitemgubun %></a></td>
        <td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipLedge.FItemList(i).getMeaipTypeName %></a></td>
        <td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipLedge.FItemList(i).FstockPlace %></a></td>
		<% if (showShopid <> "") then %>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipLedge.FItemList(i).Fshopid %></a></td>
		<% end if %>
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

		<td><a href="<%= iURLEtc %>&chulgogubun=etc" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo,0) %></a></td>
		<td><a href="<%= iURLEtc %>&chulgogubun=etc" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum,0) %></a></td>

		<td><a href="<%= iURLEtc %>&chulgogubun=loss" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulNo,0) %></a></td>
		<td><a href="<%= iURLEtc %>&chulgogubun=loss" target="_blank"><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FLossChulSum,0) %></a></td>

		<td>
			<% if (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun = "10") and (oCMonthlyMaeipLedge.FItemList(i).FstockPlace = "L") then %>
			<a href="/cscenter/action/cschulgolist.asp?menupos=1768&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>" target="_blank">
			<% end if %>
			<%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FCsNo,0) %>
		</td>
		<td>
			<% if (oCMonthlyMaeipLedge.FItemList(i).Fitemgubun = "10") and (oCMonthlyMaeipLedge.FItemList(i).FstockPlace = "L") then %>
			<a href="/cscenter/action/cschulgolist.asp?menupos=1768&yyyy1=<%= yyyy1 %>&mm1=<%= mm1 %>" target="_blank">
			<% end if %>
			<%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FCsSum,0) %>
		</td>

		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getTotErrNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).getTotErrSum,0) %></td>

		<td><b><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo,0) %></b></td>
		<td><b><%= FormatNumber(oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum,0) %></b></td>
    </tr>
	<% end if %>
	<% next %>

    <tr align="center" bgcolor="#EEEEEE">
    	<td colspan="<%= itemgubunColNUM %>">저장품소계</td>
    	<td align="right" ><%= FormatNumber(totprevSysStockNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totprevSysStockSum2,0) %></td>

		<td align="right" ><%= FormatNumber(totIpgoNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totIpgoSum2,0) %></td>

        <td align="right" ><%= FormatNumber(totMoveNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totMoveSum2,0) %></td>

		<td align="right" ><%= FormatNumber(totSellNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totSellSum2,0) %></td>

		<td align="right" ><%= FormatNumber(totOffChulNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totOffChulSum2,0) %></td>

		<td align="right" ><%= FormatNumber(totEtcChulNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totEtcChulSum2,0) %></td>

		<td align="right" ><%= FormatNumber(totLossChulNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totLossChulSum2,0) %></td>

		<td align="right" ><%= FormatNumber(totCsNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totCsSum2,0) %></td>

		<td align="right" ><%= FormatNumber(totErrNo2,0) %></td>
		<td align="right" ><%= FormatNumber(totErrSum2,0) %></td>

		<td align="right" ><b><%= FormatNumber(totcurSysStockNo2,0) %></b></td>
		<td align="right" ><b><%= FormatNumber(totcurSysStockSum2,0) %></b></td>
    </tr>

    <tr align="center" bgcolor="#FFFFFF">
    	<td colspan="<%= itemgubunColNUM %>">합계</td>
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

		<td align="right" ><b><%= FormatNumber(totcurSysStockNo,0) %></b></td>
		<td align="right" ><b><%= FormatNumber(totcurSysStockSum,0) %></b></td>
    </tr>

	<% if (Left(Now(),7) < "2016-01") then %><!-- 정산내역 숨기기 -->

    <tr align="center" bgcolor="#FFFFFF"height="25">
        <td colspan="26" align="left">[정산내역]</td>
    </tr>

    <% for i=0 to oCMonthlyMaeipJungsan.FResultCount-1 %>
    <%

	jtotprevSysStockNo       	= jtotprevSysStockNo + oCMonthlyMaeipJungsan.FItemList(i).FprevSysStockNo
	jtotprevSysStockSum       	= jtotprevSysStockSum + oCMonthlyMaeipJungsan.FItemList(i).FprevSysStockSum

	jtotIpgoNo       			= jtotIpgoNo + oCMonthlyMaeipJungsan.FItemList(i).getIpgoNo
	jtotIpgoSum       			= jtotIpgoSum + oCMonthlyMaeipJungsan.FItemList(i).getIpgoSum

    jtotMoveNo       			= jtotMoveNo + oCMonthlyMaeipJungsan.FItemList(i).getMoveNo
    jtotMoveSum       			= jtotMoveSum + oCMonthlyMaeipJungsan.FItemList(i).getMoveSum

	jtotSellNo       			= jtotSellNo + oCMonthlyMaeipJungsan.FItemList(i).FSellNo
	jtotSellSum       			= jtotSellSum + oCMonthlyMaeipJungsan.FItemList(i).FSellSum

	jtotOffChulNo       			= jtotOffChulNo + oCMonthlyMaeipJungsan.FItemList(i).FOffChulNo
	jtotOffChulSum       		= jtotOffChulSum + oCMonthlyMaeipJungsan.FItemList(i).FOffChulSum

	jtotEtcChulNo       			= jtotEtcChulNo + oCMonthlyMaeipJungsan.FItemList(i).FEtcChulNo
	jtotEtcChulSum       		= jtotEtcChulSum + oCMonthlyMaeipJungsan.FItemList(i).FEtcChulSum

	jtotLossChulNo       		= jtotLossChulNo + oCMonthlyMaeipJungsan.FItemList(i).FLossChulNo
	jtotLossChulSum       		= jtotLossChulSum + oCMonthlyMaeipJungsan.FItemList(i).FLossChulSum

	jtotCsNo       				= jtotCsNo + oCMonthlyMaeipJungsan.FItemList(i).FCsNo
	jtotCsSum       				= jtotCsSum + oCMonthlyMaeipJungsan.FItemList(i).FCsSum

	jtotcurSysStockNo       		= jtotcurSysStockNo + oCMonthlyMaeipJungsan.FItemList(i).FcurSysStockNo
	jtotcurSysStockSum       	= jtotcurSysStockSum + oCMonthlyMaeipJungsan.FItemList(i).FcurSysStockSum
	jtotcurErrRealCheckNo       	= jtotcurErrRealCheckNo + oCMonthlyMaeipJungsan.FItemList(i).FcurErrRealCheckNo
	jtotcurErrRealCheckSum       = jtotcurErrRealCheckSum + oCMonthlyMaeipJungsan.FItemList(i).FcurErrRealCheckSum

	'jdiff = oCMonthlyMaeipJungsan.FItemList(i).FprevSysStockNo + oCMonthlyMaeipJungsan.FItemList(i).getIpgoNo + oCMonthlyMaeipJungsan.FItemList(i).getMoveNo + oCMonthlyMaeipJungsan.FItemList(i).FSellNo + oCMonthlyMaeipJungsan.FItemList(i).FOffChulNo + oCMonthlyMaeipJungsan.FItemList(i).FEtcChulNo + oCMonthlyMaeipJungsan.FItemList(i).FCsNo + oCMonthlyMaeipJungsan.FItemList(i).FLossChulNo - oCMonthlyMaeipJungsan.FItemList(i).FcurSysStockNo
	'jtotdiff = jtotdiff + jdiff

	iURL = "monthlyMaeipLedge_summaryDetail.asp?menupos="& menupos
	iURL = iURL + "&yyyy1="& yyyy1 &"&mm1="& mm1 &"&makerid="& makerid &"&showsuply="&showsuply&"&meaipTp="&oCMonthlyMaeipJungsan.FItemList(i).Flastmwdiv
    iURL = iURL + "&itemgubun="&oCMonthlyMaeipJungsan.FItemList(i).Fitemgubun&"&targetGbn="&oCMonthlyMaeipJungsan.FItemList(i).FtargetGbn
    iURL = iURL + "&stockPlace="&oCMonthlyMaeipJungsan.FItemList(i).FstockPlace&"&shopid="&oCMonthlyMaeipJungsan.FItemList(i).Fshopid&"&stype=J"
    %>
    <tr align="right" bgcolor="#FFFFFF" >

		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipJungsan.FItemList(i).getITemGubunName%></a></td>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipJungsan.FItemList(i).Fitemgubun %></a></td>
        <td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipJungsan.FItemList(i).getMeaipTypeName %></a></td>
        <td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipJungsan.FItemList(i).Fshopid %></a></td>
		<% if (showShopid <> "") then %>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyMaeipJungsan.FItemList(i).Fshopid %></a></td>
		<% end if %>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FprevSysStockNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FprevSysStockSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).getIpgoNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).getIpgoSum,0) %></td>

        <td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).getMoveNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).getMoveSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FSellNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FSellSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FOffChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FOffChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FEtcChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FEtcChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FLossChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FLossChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FCsNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FCsSum,0) %></td>

		<td></td>
		<td></td>

		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FcurSysStockNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyMaeipJungsan.FItemList(i).FcurSysStockSum,0) %></td>


    </tr>
	<% next %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
		<td></td>
    	<td></td>
        <td></td>
		<% if (showShopid <> "") then %>
		<td></td>
		<% end if %>
    	<td align="right" ><%= FormatNumber(jtotprevSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotprevSysStockSum,0) %></td>

		<td align="right" ><%= FormatNumber(jtotIpgoNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotIpgoSum,0) %></td>

		<td align="right" ><%= FormatNumber(jtotMoveNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotMoveSum,0) %></td>

		<td align="right" ><%= FormatNumber(jtotSellNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotSellSum,0) %></td>

		<td align="right" ><%= FormatNumber(jtotOffChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotOffChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(jtotEtcChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotEtcChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(jtotLossChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotLossChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(jtotCsNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotCsSum,0) %></td>

		<td align="right" ></td>
		<td align="right" ></td>

		<td align="right" ><%= FormatNumber(jtotcurSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(jtotcurSysStockSum,0) %></td>


    </tr>

    <tr align="center" bgcolor="#CCCCCC">
    	<td></td>
		<td></td>
    	<td></td>
        <td></td>
		<% if (showShopid <> "") then %>
		<td></td>
		<% end if %>
    	<td align="right" ><%= FormatNumber(totprevSysStockNo+jtotprevSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(totprevSysStockSum+jtotprevSysStockSum,0) %></td>

		<td align="right" ><%= FormatNumber(totIpgoNo+jtotIpgoNo,0) %></td>
		<td align="right" ><%= FormatNumber(totIpgoSum+jtotIpgoSum,0) %></td>

		<td align="right" ><%= FormatNumber(totMoveNo+jtotMoveNo,0) %></td>
		<td align="right" ><%= FormatNumber(totMoveSum+jtotMoveSum,0) %></td>

		<td align="right" ><%= FormatNumber(totSellNo+jtotSellNo,0) %></td>
		<td align="right" ><%= FormatNumber(totSellSum+jtotSellSum,0) %></td>

		<td align="right" ><%= FormatNumber(totOffChulNo+jtotOffChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totOffChulSum+jtotOffChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totEtcChulNo+jtotEtcChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totEtcChulSum+jtotEtcChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totLossChulNo+jtotLossChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(totLossChulSum+jtotLossChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(totCsNo+jtotCsNo,0) %></td>
		<td align="right" ><%= FormatNumber(totCsSum+jtotCsSum,0) %></td>

		<td align="right" ></td>
		<td align="right" ></td>

        <td align="right" ><%= FormatNumber(totcurSysStockNo+jtotcurSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(totcurSysStockSum+jtotcurSysStockSum,0) %></td>


    </tr>
    <% if (oCMonthlyIpMaeipJungsan.FresultCount>0) then %>
    <tr align="center" bgcolor="#FFFFFF"height="25">
        <td colspan="26" align="left">[입고분 정산]</td>
    </tr>

    <% for i=0 to oCMonthlyIpMaeipJungsan.FResultCount-1 %>
    <%

	jIptotprevSysStockNo       	= jIptotprevSysStockNo + oCMonthlyIpMaeipJungsan.FItemList(i).FprevSysStockNo
	jIptotprevSysStockSum       	= jIptotprevSysStockSum + oCMonthlyIpMaeipJungsan.FItemList(i).FprevSysStockSum

	jIptotIpgoNo       			= jIptotIpgoNo + oCMonthlyIpMaeipJungsan.FItemList(i).getIpgoNo
	jIptotIpgoSum       			= jIptotIpgoSum + oCMonthlyIpMaeipJungsan.FItemList(i).getIpgoSum

    jIptotMoveNo       			= jIptotMoveNo + oCMonthlyIpMaeipJungsan.FItemList(i).getMoveNo
    jIptotMoveSum       			= jIptotMoveSum + oCMonthlyIpMaeipJungsan.FItemList(i).getMoveSum


	jIptotSellNo       			= jIptotSellNo + oCMonthlyIpMaeipJungsan.FItemList(i).FSellNo
	jIptotSellSum       			= jIptotSellSum + oCMonthlyIpMaeipJungsan.FItemList(i).FSellSum

	jIptotOffChulNo       			= jIptotOffChulNo + oCMonthlyIpMaeipJungsan.FItemList(i).FOffChulNo
	jIptotOffChulSum       		= jIptotOffChulSum + oCMonthlyIpMaeipJungsan.FItemList(i).FOffChulSum

	jIptotEtcChulNo       			= jIptotEtcChulNo + oCMonthlyIpMaeipJungsan.FItemList(i).FEtcChulNo
	jIptotEtcChulSum       		= jIptotEtcChulSum + oCMonthlyIpMaeipJungsan.FItemList(i).FEtcChulSum

	jIptotLossChulNo       		= jIptotLossChulNo + oCMonthlyIpMaeipJungsan.FItemList(i).FLossChulNo
	jIptotLossChulSum       		= jIptotLossChulSum + oCMonthlyIpMaeipJungsan.FItemList(i).FLossChulSum

	jIptotCsNo       				= jIptotCsNo + oCMonthlyIpMaeipJungsan.FItemList(i).FCsNo
	jIptotCsSum       				= jIptotCsSum + oCMonthlyIpMaeipJungsan.FItemList(i).FCsSum

	jIptotcurSysStockNo       		= jIptotcurSysStockNo + oCMonthlyIpMaeipJungsan.FItemList(i).FcurSysStockNo
	jIptotcurSysStockSum       	= jIptotcurSysStockSum + oCMonthlyIpMaeipJungsan.FItemList(i).FcurSysStockSum
	jIptotcurErrRealCheckNo       	= jIptotcurErrRealCheckNo + oCMonthlyIpMaeipJungsan.FItemList(i).FcurErrRealCheckNo
	jIptotcurErrRealCheckSum       = jIptotcurErrRealCheckSum + oCMonthlyIpMaeipJungsan.FItemList(i).FcurErrRealCheckSum

	'jIpdiff = oCMonthlyIpMaeipJungsan.FItemList(i).FprevSysStockNo + oCMonthlyIpMaeipJungsan.FItemList(i).getIpgoNo + oCMonthlyIpMaeipJungsan.FItemList(i).getMoveNo + oCMonthlyIpMaeipJungsan.FItemList(i).FSellNo + oCMonthlyIpMaeipJungsan.FItemList(i).FOffChulNo + oCMonthlyIpMaeipJungsan.FItemList(i).FEtcChulNo + oCMonthlyIpMaeipJungsan.FItemList(i).FCsNo + oCMonthlyIpMaeipJungsan.FItemList(i).FLossChulNo - oCMonthlyIpMaeipJungsan.FItemList(i).FcurSysStockNo
	'jIptotdiff = jIptotdiff + jIpdiff

	iURL = "monthlyMaeipLedge_summaryDetail.asp?menupos="& menupos
	iURL = iURL + "&yyyy1="& yyyy1 &"&mm1="& mm1 &"&makerid="& makerid &"&showsuply="&showsuply&"&meaipTp="&oCMonthlyIpMaeipJungsan.FItemList(i).Flastmwdiv
    iURL = iURL + "&itemgubun="&oCMonthlyIpMaeipJungsan.FItemList(i).Fitemgubun&"&targetGbn="&oCMonthlyIpMaeipJungsan.FItemList(i).FtargetGbn
    iURL = iURL + "&stockPlace="&oCMonthlyIpMaeipJungsan.FItemList(i).FstockPlace&"&shopid="&oCMonthlyIpMaeipJungsan.FItemList(i).Fshopid&"&stype=J"
    %>
    <tr align="right" bgcolor="#FFFFFF" >

		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyIpMaeipJungsan.FItemList(i).getITemGubunName%></a></td>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyIpMaeipJungsan.FItemList(i).Fitemgubun %></a></td>
        <td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyIpMaeipJungsan.FItemList(i).getMeaipTypeName %></a></td>
        <td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyIpMaeipJungsan.FItemList(i).Fshopid %></a></td>
		<% if (showShopid <> "") then %>
		<td align="center"><a href="<%= iURL %>" target="_blank"><%= oCMonthlyIpMaeipJungsan.FItemList(i).Fshopid %></a></td>
		<% end if %>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FprevSysStockNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FprevSysStockSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).getIpgoNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).getIpgoSum,0) %></td>

        <td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).getMoveNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).getMoveSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FSellNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FSellSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FOffChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FOffChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FEtcChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FEtcChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FLossChulNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FLossChulSum,0) %></td>

		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FCsNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FCsSum,0) %></td>

        <td></td>
		<td></td>

		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FcurSysStockNo,0) %></td>
		<td><%= FormatNumber(oCMonthlyIpMaeipJungsan.FItemList(i).FcurSysStockSum,0) %></td>


    </tr>
	<% next %>

    <tr align="center" bgcolor="#FFFFFF">
    	<td></td>
		<td></td>
    	<td></td>
        <td></td>
		<% if (showShopid <> "") then %>
		<td></td>
		<% end if %>
    	<td align="right" ><%= FormatNumber(jIptotprevSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotprevSysStockSum,0) %></td>

		<td align="right" ><%= FormatNumber(jIptotIpgoNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotIpgoSum,0) %></td>

		<td align="right" ><%= FormatNumber(jIptotMoveNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotMoveSum,0) %></td>

		<td align="right" ><%= FormatNumber(jIptotSellNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotSellSum,0) %></td>

		<td align="right" ><%= FormatNumber(jIptotOffChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotOffChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(jIptotEtcChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotEtcChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(jIptotLossChulNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotLossChulSum,0) %></td>

		<td align="right" ><%= FormatNumber(jIptotCsNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotCsSum,0) %></td>

		<td></td>
		<td></td>

		<td align="right" ><%= FormatNumber(jIptotcurSysStockNo,0) %></td>
		<td align="right" ><%= FormatNumber(jIptotcurSysStockSum,0) %></td>


    </tr>
    <% end if %>

	<% end if %><!-- 정산내역 숨기기 -->

</table>

<% IF application("Svr_Info")="Dev" THEN %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="300"></iframe>
<% else %>
	<iframe src="about:blank" name="exceldown" border="0" width="100%" height="0"></iframe>
<% end if %>

<%
set oCMonthlyMaeipLedge = Nothing
set oCMonthlyMaeipJungsan = Nothing
set oCMonthlyIpMaeipJungsan = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
