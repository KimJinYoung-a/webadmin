<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  핑거스 매출집계-일별
' History : 2016.09.20 한용민 생성
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/report/maechul/statisticCls.asp" -->
<%
dim menupos : menupos = getNumeric(requestcheckvar(request("menupos"),10))
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay
Dim vTot_CountPlus, vTot_CountMinus, vTot_MaechulPlus, vTot_MaechulMinus, vTot_Subtotalprice, vTot_Miletotalprice, vTot_subtotalprice_notexists_sumPaymentEtc
dim vTot_MaechulCountSum, vTot_MaechulPriceSum, vTot_sumPaymentEtc, page, pagesize, vSorting
dim sellchnl
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= RequestCheckvar(request("sitename"),16)
	vDateGijun	= NullFillWith(RequestCheckvar(request("date_gijun"),16),"regdate")
	vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(DateAdd("d",-13,now())))
	vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(DateAdd("d",-13,now())))
	vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),Day(DateAdd("d",-13,now())))
	vEYear		= NullFillWith(RequestCheckvar(request("eyear"),4),Year(now))
	vEMonth		= NullFillWith(RequestCheckvar(request("emonth"),2),Month(now))
	vEDay		= NullFillWith(RequestCheckvar(request("eday"),2),Day(now))
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"ddateD")

if (page = "") then
	page = 1
end if

pagesize = 5000

Set cStatistic = New cacademyStatic_list
	cStatistic.FCurrPage = page
	cStatistic.FPageSize = pagesize
	cStatistic.FRectSort = vSorting
	cStatistic.FRectDateGijun = vDateGijun
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.FRectSellChannelDiv = sellchnl
	cStatistic.facademyStatistic_Sexdailylist()

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF">
	<td colspan="25">
		검색결과 : <b><%=cStatistic.FresultCount%></b>&nbsp;&nbsp;※ 최대 5천건 까지 보여 집니다.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td rowspan="2" colspan="2">
		기간
	</td>
    <td colspan="2">여자 매출액</td>
    <td colspan="2">남자 매출액</td>
    <td colspan="2">매출액합계</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>
    	주문건수
    </td>
    <td>
    	금액
    </td>
    <td>
    	주문건수
    </td>
    <td>
    	금액
    </td>
    <td>
    	주문건수
    </td>
    <td>
    	금액
    </td>
</tr>

<% if cStatistic.FTotalCount > 0 then %>
	<% For i = 0 To cStatistic.FTotalCount -1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center">
			<% if right(FormatDateTime(cStatistic.FItemList(i).FRegdate,1),3) = "토요일" then %>
				<font color="blue"><%= cStatistic.FItemList(i).FRegdate %></font>
			<% elseif right(FormatDateTime(cStatistic.FItemList(i).FRegdate,1),3) = "일요일" then %>
				<font color="red"><%= cStatistic.FItemList(i).FRegdate %></font>
			<% else %>
				<%= cStatistic.FItemList(i).FRegdate %>
			<% end if %>
		</td>
		<td align="center"><%= getWeekdayStr(DatePart("w",cStatistic.FItemList(i).FRegdate)) %></td>
		<td align="center"><%= FormatNumber(cStatistic.FItemList(i).FCountMinus,0) %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FMaechulMinus,0) %></td>
		<td align="center"><%= FormatNumber(cStatistic.FItemList(i).FCountPlus,0) %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FMaechulPlus,0) %></td>
		<td align="center"><%= FormatNumber(cStatistic.FItemList(i).fcount_plus_minus) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><b><%= FormatNumber(cStatistic.FItemList(i).fmaechul_plus_minus,0) %></b></td>
	</tr>
	<%
	vTot_CountPlus			= vTot_CountPlus + CLng(FormatNumber(cStatistic.FItemList(i).FCountPlus,0))
	vTot_MaechulPlus		= vTot_MaechulPlus + CLng(FormatNumber(cStatistic.FItemList(i).FMaechulPlus,0))
	vTot_CountMinus			= vTot_CountMinus + CLng(FormatNumber(cStatistic.FItemList(i).FCountMinus,0))
	vTot_MaechulMinus		= vTot_MaechulMinus + CLng(FormatNumber(cStatistic.FItemList(i).FMaechulMinus,0))
	vTot_MaechulCountSum	= vTot_MaechulCountSum + CLng(FormatNumber(cStatistic.FItemList(i).fcount_plus_minus,0))
	vTot_MaechulPriceSum	= vTot_MaechulPriceSum + CLng(FormatNumber(cStatistic.FItemList(i).fmaechul_plus_minus,0))
	Next
	%>
	<tr bgcolor="<%= adminColor("tabletop") %>">
		<td align="center" colspan="2">합계</td>
		<td align="center"><%=FormatNumber(vTot_CountMinus)%></td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulMinus,0)%></td>
		<td align="center"><%=FormatNumber(vTot_CountPlus)%></td>
		<td align="right" style="padding-right:5px;"><%=FormatNumber(vTot_MaechulPlus,0)%></td>
		<td align="center"><%=FormatNumber(vTot_MaechulCountSum)%></td>
		<td align="right" style="padding-right:5px;"><b><%=FormatNumber(vTot_MaechulPriceSum,0)%></b></td>
	</tr>
<% ELSE %>
	<tr  align="center" bgcolor="#FFFFFF">
		<td colspan="25">등록된 내용이 없습니다.</td>
	</tr>
<% end if %>

</table>

<%
Set cStatistic = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->