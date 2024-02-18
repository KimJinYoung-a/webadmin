<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인 매출통계 월결제유형별통계 실시간
' History : 2012.11.05 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->

<%
Dim i, cStatistic, shopid, datefg, yyyy1, mm1, stdate, BanPum, fromDate, toDate, offgubun, reload, inc3pl
dim vTot_spendmile, vTot_TenGiftCardPaySum, vTot_giftcardPaysum, vTot_cardsum, vTot_cashsum, vTot_TotalSum, vTot_extPaySum
	shopid 	= requestCheckVar(request("shopid"),32)
	datefg = requestCheckVar(request("datefg"),32)
	yyyy1 = requestCheckVar(request("yyyy1"),4)
	mm1	  = requestCheckVar(request("mm1"),2)
	BanPum     = requestCheckVar(request("BanPum"),1)
	offgubun = requestCheckVar(request("offgubun"),10)
	reload = requestCheckVar(request("reload"),2)
    inc3pl = requestCheckVar(request("inc3pl"),32)

if reload <> "on" and offgubun = "" then offgubun = "95"
if datefg = "" then datefg = "maechul"
if yyyy1="" then
	stdate = CStr(Now)
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2)),1)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if

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
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''다른매장조회 막음.
		else
		end if
	end if
end if

Set cStatistic = New COffShopSellReport
	cStatistic.FRectdatefg = datefg
	cStatistic.FRectStartdate = yyyy1 + "-" + mm1 + "-" + "01"
	cStatistic.FRectEnddate = CStr(DateAdd("m",1,DateSerial(yyyy1,mm1,1)))
	cStatistic.FRectshopid = shopid
	cStatistic.FRectBanPum = BanPum
	cStatistic.FRectOffgubun = offgubun
	cStatistic.FRectInc3pl = inc3pl
	cStatistic.FPageSize = 500
	cStatistic.FCurrPage = 1
	cStatistic.GetJumunMethodReportMonth()
%>

<script language="javascript">

function searchSubmit(){
	frm.submit();
}

</script>

<!-- 검색 시작 -->
<form name="frm" method="get" style="margin:0px;">
	<input type="hidden" name="reload" value="on">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center" bgcolor="#FFFFFF" >
	<td width="30" bgcolor="<%= adminColor("gray") %>">검색<Br>조건</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* 기간 :&nbsp;
				<% drawmaechul_datefg "datefg" ,datefg ,""%>
				<% DrawYMBox yyyy1,mm1 %>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* 매장 : <% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
					<% else %>
						<!--* 매장 : <%' drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='searchSubmit();'","" %>-->
					<% end if %>
				<% end if %>
				&nbsp;&nbsp;
				* 매장 구분 : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='searchSubmit();'" %>
				<Br>
				* 반품여부 :
				<% drawSelectBoxisusingYN "BanPum" , BanPum ," onchange='searchSubmit();'" %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="검색" onClick="javascript:searchSubmit();"></td>
		</tr>
	</table>
</form>
<!-- 검색 끝 -->
<br>
<!-- 표 중간바 시작-->
<table width="100%" cellpadding="1" cellspacing="1" class="a">
	<tr valign="bottom">
    <td align="left">
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=cStatistic.FTotalCount%></b> ※ 총 500건까지 검색됩니다.
	</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">기간</td>
    <td align="center" colspan="3"></td>
    <td align="center" colspan="3">실결제액</td>
    <td align="center" width="150" rowspan="2">매출합계</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">마일리지</td>
    <td align="center">기프트카드</td>
    <td align="center">상품권</td>
    <td align="center">신용카드</td>
    <td align="center">현금</td>
	<td align="center">기타</td>
	</tr>
<%
if cStatistic.FTotalCount > 0 then

For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF">
	<td align="center">
		<%= getweekendcolor(cStatistic.fitemlist(i).FRegdate) %>
	</td>
	<td align="center"><%= getweekend(cStatistic.fitemlist(i).FRegdate) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fspendmile,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fTenGiftCardPaySum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fgiftcardPaysum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcardsum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcashsum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fextPaysum,0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><%= FormatNumber(cStatistic.FItemList(i).fselltotal,0) %></td>
	</tr>
<%
	vTot_spendmile	= vTot_spendmile + CLng(cStatistic.FItemList(i).fspendmile)
	vTot_TenGiftCardPaySum		= vTot_TenGiftCardPaySum + CLng(cStatistic.FItemList(i).fTenGiftCardPaySum)
	vTot_giftcardPaysum		= vTot_giftcardPaysum + CLng(cStatistic.FItemList(i).fgiftcardPaysum)
	vTot_extPaysum		= vTot_extPaysum + CLng(cStatistic.FItemList(i).fextPaysum)
	vTot_cardsum		= vTot_cardsum + CLng(cStatistic.FItemList(i).fcardsum)
	vTot_cashsum			= vTot_cashsum + CLng(cStatistic.FItemList(i).fcashsum)
	vTot_TotalSum		= vTot_TotalSum + CLng(cStatistic.FItemList(i).fselltotal)

Next
%>
	<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2">합계</td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_spendmile,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_TenGiftCardPaySum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_giftcardPaysum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_cardsum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_cashsum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_extPaysum,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_TotalSum,0) %></td>
	</tr>
<% else %>
	<tr align="center" bgcolor="#FFFFFF">
	<td colspan="25">등록된 내용이 없습니다.</td>
	</tr>
<% end if %>
</table>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
