<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  오프라인 매출통계 일별
' History : 2012.10.04 한용민 생성
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
Dim i, cStatistic, shopid, datefg,yyyy1,mm1,dd1,yyyy2,mm2,dd2, offgubun, reload, oldlist
Dim vTot_CountPlus, vTot_CountMinus, vTot_MaechulPlus, vTot_MaechulMinus, vTot_Subtotalprice
dim vTot_Miletotalprice, vTot_MaechulCountSum, vTot_MaechulPriceSum ,fromDate ,toDate, inc3pl
dim xl
	shopid 	= request("shopid")
	datefg = request("datefg")
	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")
	offgubun = request("offgubun")
	reload = request("reload")
	oldlist = request("oldlist")
    inc3pl = request("inc3pl")
	xl 			= request("xl")

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
	cStatistic.FRectOffgubun = offgubun
	cStatistic.FRectOldData = oldlist
	cStatistic.FRectInc3pl = inc3pl
	cStatistic.fStatistic_dailylist_datamart()

if (xl = "Y") then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=datamart_off_daily_xl.xls"
else

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

function popXL()
{
    frmXL.submit();
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
				* 기간 :
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
						* 매장 : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,11", "", " onchange='searchSubmit();'" %>
					<% end if %>
				<% else %>
					<% if not(C_IS_Maker_Upche) then %>
						* 매장 : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,11", "", " onchange='searchSubmit();'" %>
					<% else %>
						<!--* 매장 : <%' drawBoxDirectIpchulOffShopByMakerchfg "shopid",shopid,makerid," onchange='searchSubmit();'","" %>-->
					<% end if %>
				<% end if %>
				<br>
				* 매장 구분 : <% drawoffshop_commoncode "offgubun", offgubun, "shopdivithinkso", "", "", " onchange='searchSubmit();'" %>
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

<p />

<!-- 표 중간바 시작-->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="#EEEEEE">
	<tr>
		<td align="left">
			※ 검색 기간이 길면 느려집니다. 검색 버튼을 누른뒤, 아무 반응이 없어보인다고, 다시 검색버튼을 클릭하지 마세요.
		</td>
		<td align="right">
			<input type="button" class="button" value="엑셀받기" onClick="popXL()">
		</td>
	</tr>
</table>
<!-- 표 중간바 끝-->

<p />

<% end if %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=cStatistic.FTotalCount%></b>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" rowspan="2" colspan="2">기간</td>
    <td align="center" colspan="2">매출액(+)</td>
    <td align="center" colspan="2">매출액(-)</td>
    <td align="center" colspan="2">매출액합계</td>
    <td align="center" width="150" rowspan="2">마일리지</td>
    <td align="center" width="150" rowspan="2">결제총액</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>">
    <td align="center">주문건수</td>
    <td align="center">금액</td>
    <td align="center">주문건수</td>
    <td align="center">금액</td>
    <td align="center">주문건수</td>
    <td align="center">금액</td>
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
	<td align="center"><%= FormatNumber(cStatistic.fitemlist(i).FCountPlus,0) %></td>
	<td align="right" style="padding-right:10px;"><%= FormatNumber(cStatistic.fitemlist(i).FMaechulPlus,0) %></td>
	<td align="center"><%= FormatNumber(cStatistic.fitemlist(i).FCountMinus,0) %></td>
	<td align="right" style="padding-right:10px;"><%= FormatNumber(cStatistic.fitemlist(i).FMaechulMinus,0) %></td>
	<td align="center"><%= FormatNumber(CLng(cStatistic.fitemlist(i).FCountPlus)+CLng(cStatistic.fitemlist(i).FCountMinus),0) %></td>
	<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8">
		<%= FormatNumber(CLng(cStatistic.fitemlist(i).FMaechulPlus)+CLng(cStatistic.fitemlist(i).FMaechulMinus),0) %>
	</td>
	<td align="center"><%= FormatNumber(cStatistic.fitemlist(i).FMiletotalprice,0) %></td>
	<td align="right" style="padding-right:5px;">
		<%= FormatNumber(cStatistic.fitemlist(i).FSubtotalprice,0) %>
	</td>
</tr>
<%
vTot_CountPlus			= vTot_CountPlus + cStatistic.fitemlist(i).FCountPlus
vTot_MaechulPlus		= vTot_MaechulPlus + cStatistic.fitemlist(i).FMaechulPlus
vTot_CountMinus			= vTot_CountMinus + cStatistic.fitemlist(i).FCountMinus
vTot_MaechulMinus		= vTot_MaechulMinus + cStatistic.fitemlist(i).FMaechulMinus
vTot_MaechulCountSum	= vTot_MaechulCountSum + (CLng(cStatistic.fitemlist(i).FCountPlus)+CLng(cStatistic.fitemlist(i).FCountMinus))
vTot_MaechulPriceSum	= vTot_MaechulPriceSum + (CLng(cStatistic.fitemlist(i).FMaechulPlus)+CLng(cStatistic.fitemlist(i).FMaechulMinus))
vTot_Miletotalprice		= vTot_Miletotalprice + cStatistic.fitemlist(i).FMiletotalprice
vTot_Subtotalprice		= vTot_Subtotalprice + cStatistic.fitemlist(i).FSubtotalprice

Next
%>
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="center" colspan="2">합계</td>
	<td align="center"><%=FormatNumber(vTot_CountPlus,0)%></td>
	<td align="right" style="padding-right:10px;"><%=FormatNumber(vTot_MaechulPlus,0)%></td>
	<td align="center"><%=FormatNumber(vTot_CountMinus,0)%></td>
	<td align="right" style="padding-right:10px;"><%=FormatNumber(vTot_MaechulMinus,0)%></td>
	<td align="center"><%=FormatNumber(vTot_MaechulCountSum,0)%></td>
	<td align="right" style="padding-right:10px;"><%=FormatNumber(vTot_MaechulPriceSum,0)%></td>
	<td align="center"><%=FormatNumber(vTot_Miletotalprice,0)%></td>
	<td align="right" style="padding-right:10px;"><%=FormatNumber(vTot_Subtotalprice,0)%></td>
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

<form name="frmXL" method="get" style="margin:0px;">
	<input type="hidden" name="xl" value="Y">
	<input type="hidden" name="shopid" value="<%= shopid %>">
	<input type="hidden" name="datefg" value="<%= datefg %>">
	<input type="hidden" name="yyyy1" value="<%= yyyy1 %>">
	<input type="hidden" name="mm1" value="<%= mm1 %>">
	<input type="hidden" name="dd1" value="<%= dd1 %>">
	<input type="hidden" name="yyyy2" value="<%= yyyy2 %>">
	<input type="hidden" name="mm2" value="<%= mm2 %>">
	<input type="hidden" name="dd2" value="<%= dd2 %>">
	<input type="hidden" name="offgubun" value="<%= offgubun %>">
	<input type="hidden" name="reload" value="<%= reload %>">
	<input type="hidden" name="oldlist" value="<%= oldlist %>">
	<input type="hidden" name="inc3pl" value="<%= inc3pl %>">
</form>

<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
