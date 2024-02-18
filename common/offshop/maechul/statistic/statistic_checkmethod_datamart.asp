<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인 매출통계 매출집계-결제방식별
' History : 2012.11.05 한용민 생성
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/maechul/statistic/statisticCls_datamart.asp" -->

<%
Dim i, cStatistic, shopid, datefg,yyyy1,mm1,dd1,yyyy2,mm2,dd2 ,BanPum ,fromDate ,toDate, offgubun, reload
dim vTot_spendmile, vTot_TenGiftCardPaySum, vTot_giftcardPaysum, vTot_cardsum, vTot_cashsum, vTot_TotalSum
dim oldlist, vTot_spendmilecnt, vTot_TenGiftCardPaycount, vTot_giftcardPaycnt, vTot_cardcnt, vTot_cashcnt
dim inc3pl
dim onlyTenShop
	shopid 	= request("shopid")
	datefg = request("datefg")
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	BanPum     = request("BanPum")
	offgubun = request("offgubun")
	reload = request("reload")
	oldlist = request("oldlist")
    inc3pl = request("inc3pl")
	onlyTenShop = request("onlyTenShop")

if reload <> "on" and offgubun = "" then offgubun = "95"
if datefg = "" then datefg = "maechul"
if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

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

Set cStatistic = New cStaticdatamart_list
	cStatistic.FRectdatefg = datefg
	cStatistic.FRectStartdate = fromDate
	cStatistic.FRectEndDate = toDate
	cStatistic.FRectshopid = shopid
	cStatistic.FRectBanPum = BanPum
	cStatistic.FRectOffgubun = offgubun
	cStatistic.FRectOldData = oldlist
	cStatistic.FRectInc3pl = inc3pl

	cStatistic.FRectOnlyTenShop = onlyTenShop

	cStatistic.fStatistic_checkmethod_datamart()

vTot_spendmile=0
vTot_TenGiftCardPaySum=0
vTot_giftcardPaysum=0
vTot_cardsum=0
vTot_cashsum=0
vTot_TotalSum=0
vTot_spendmilecnt=0
vTot_TenGiftCardPaycount=0
vTot_giftcardPaycnt=0
vTot_cardcnt=0
vTot_cashcnt=0
%>

<script language="javascript">

function searchSubmit()
{
	//날짜 비교
	var startdate = frm.yyyy1.value + "-" + frm.mm1.value + "-" + frm.dd1.value;
	var enddate = frm.yyyy2.value + "-" + frm.mm2.value + "-" + frm.dd2.value;
    var diffDay = 0;
    var start_yyyy = startdate.substring(0,4);
    var start_mm = startdate.substring(5,7);
    var start_dd = startdate.substring(8,startdate.length);
    var sDate = new Date(start_yyyy, start_mm-1, start_dd);
    var end_yyyy = enddate.substring(0,4);
    var end_mm = enddate.substring(5,7);
    var end_dd = enddate.substring(8,enddate.length);
    var eDate = new Date(end_yyyy, end_mm-1, end_dd);

    diffDay = Math.ceil((eDate.getTime() - sDate.getTime())/(1000*60*60*24));

	if (diffDay > 1095 && frm.oldlist.checked == false){
		alert('3년 이전 데이터는 3년이전내역조회 를 체크하셔야 합니다');
		return;
	}

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
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				<input type="checkbox" name="oldlist" <% if oldlist="on" then response.write "checked" %> >3년이전내역조회
				&nbsp;&nbsp;
				<%
				'직영/가맹점
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") then %>
						* 매장 : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* 매장 : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,9,11", "", " onchange='searchSubmit();'" %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* 매장 : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,9,11", "", " onchange='searchSubmit();'" %>
					<% else %>
						<!--* 매장 : <%' drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='searchSubmit();'","" %>-->
					<% end if %>
				<% end if %>
				<br>
				* 매장 구분 : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='searchSubmit();'" %>
				&nbsp;&nbsp;
				* 반품여부 :
				<% drawSelectBoxisusingYN "BanPum" , BanPum ," onchange='searchSubmit();'" %>
	            &nbsp;&nbsp;
	            <b>* 매출처구분</b>
	            <% Call draw3plMeachulComboBox("inc3pl",inc3pl) %>
				&nbsp;&nbsp;
				<input type="checkbox" name="onlyTenShop" value="Y" <% if (onlyTenShop = "Y") then %>checked<% end if %> >
				텐바이텐 매장만(streetshop011, streetshop014, streetshop018)
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
		※ 검색 기간이 길면 느려집니다. 검색 버튼을 누른뒤, 아무 반응이 없어보인다고, 다시 검색버튼을 클릭하지 마세요.
    </td>
    <td align="right">
    </td>
</tr>
</table>
<!-- 표 중간바 끝-->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=cStatistic.FTotalCount%></b>
	</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td rowspan="3" colspan="2">기간</td>
    <td colspan="6"></td>
    <td colspan="4">실결제액</td>
    <td width="150" rowspan="3">매출합계</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td colspan="2">마일리지</td>
    <td colspan="2">기프트카드</td>
    <td colspan="2">상품권</td>
    <td colspan="2">신용카드</td>
    <td colspan="2">현금</td>
	</tr>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>금액</td>
    <td>건수</td>
    <td>금액</td>
    <td>건수</td>
    <td>금액</td>
    <td>건수</td>
    <td>금액</td>
    <td>건수</td>
    <td>금액</td>
    <td>건수</td>
	</tr>
<%
if cStatistic.FTotalCount > 0 then

For i = 0 To cStatistic.FTotalCount -1
%>
	<tr bgcolor="#FFFFFF" align="center">
	<td>
		<%= getweekendcolor(cStatistic.fitemlist(i).FRegdate) %>
	</td>
	<td align="center"><%= getweekend(cStatistic.fitemlist(i).FRegdate) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fspendmile,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fspendmilecnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fTenGiftCardPaySum,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fTenGiftCardPaycount,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fgiftcardPaysum,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fgiftcardPaycnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcardsum,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fcardcnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).fcashsum,0) %></td>
	<td><%= FormatNumber(cStatistic.FItemList(i).fcashcnt,0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><%= FormatNumber(cStatistic.FItemList(i).fselltotal,0) %></td>
	</tr>
<%
	vTot_spendmile	= vTot_spendmile + CLng(cStatistic.FItemList(i).fspendmile)
	vTot_TenGiftCardPaySum		= vTot_TenGiftCardPaySum + CLng(cStatistic.FItemList(i).fTenGiftCardPaySum)
	vTot_giftcardPaysum		= vTot_giftcardPaysum + CLng(cStatistic.FItemList(i).fgiftcardPaysum)
	vTot_cardsum		= vTot_cardsum + CLng(cStatistic.FItemList(i).fcardsum)
	vTot_cashsum			= vTot_cashsum + CLng(cStatistic.FItemList(i).fcashsum)
	vTot_TotalSum		= vTot_TotalSum + CLng(cStatistic.FItemList(i).fselltotal)
	vTot_spendmilecnt		= vTot_spendmilecnt + CLng(cStatistic.FItemList(i).fspendmilecnt)
	vTot_TenGiftCardPaycount		= vTot_TenGiftCardPaycount + CLng(cStatistic.FItemList(i).fTenGiftCardPaycount)
	vTot_giftcardPaycnt		= vTot_giftcardPaycnt + CLng(cStatistic.FItemList(i).fgiftcardPaycnt)
	vTot_cardcnt		= vTot_cardcnt + CLng(cStatistic.FItemList(i).fcardcnt)
	vTot_cashcnt		= vTot_cashcnt + CLng(cStatistic.FItemList(i).fcashcnt)

Next
%>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td colspan="2">합계</td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_spendmile,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_spendmilecnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_TenGiftCardPaySum,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_TenGiftCardPaycount,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_giftcardPaysum,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_giftcardPaycnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_cardsum,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_cardcnt,0) %></td>
	<td align="right" style="padding-right:5px;"><%= FormatNumber(vTot_cashsum,0) %></td>
	<td style="padding-right:5px;"><%= FormatNumber(vTot_cashcnt,0) %></td>
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
<!-- #include virtual="/lib/db/db3close.asp" -->
