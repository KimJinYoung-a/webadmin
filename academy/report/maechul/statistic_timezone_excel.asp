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
Dim i, cStatistic, vSiteName, vDateGijun, v6MonthDate, vSYear, vSMonth, vSDay, vEYear, vEMonth, vEDay
Dim vTot_CountPlus, vTot_CountMinus, vTot_MaechulPlus, vTot_MaechulMinus, vTot_Subtotalprice, vTot_Miletotalprice, vTot_subtotalprice_notexists_sumPaymentEtc
dim vTot_MaechulCountSum, vTot_MaechulPriceSum, vTot_sumPaymentEtc, page, pagesize, vSorting
dim sellchnl
	v6MonthDate	= DateAdd("m",-6,now())
	vSiteName 	= RequestCheckvar(request("sitename"),16)
	vDateGijun	= NullFillWith(RequestCheckvar(request("date_gijun"),16),"regdate")
	'vSYear		= NullFillWith(request("syear"),Year(DateAdd("d",-13,now())))
	'vSMonth		= NullFillWith(request("smonth"),Month(DateAdd("d",-13,now())))
	'vSDay		= NullFillWith(request("sday"),Day(DateAdd("d",-13,now())))
	vSYear		= NullFillWith(RequestCheckvar(request("syear"),4),Year(now()))
	vSMonth		= NullFillWith(RequestCheckvar(request("smonth"),2),Month(now()))
	vSDay		= NullFillWith(RequestCheckvar(request("sday"),2),"01")
	vEYear		= NullFillWith(RequestCheckvar(request("eyear"),4),Year(now))
	vEMonth		= NullFillWith(RequestCheckvar(request("emonth"),2),Month(now))
	vEDay		= NullFillWith(RequestCheckvar(request("eday"),2),Day(now))
	sellchnl    = requestCheckVar(request("sellchnl"),20)
	vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"ddateD")

if (page = "") then
	page = 1
end if

if (pagesize = "") then
	pagesize = 3000
end if

Set cStatistic = New cacademyStatic_list
	cStatistic.FCurrPage = page
	cStatistic.FPageSize = pagesize
	cStatistic.FRectStartdate = vSYear & "-" & TwoNumber(vSMonth) & "-" & TwoNumber(vSDay)
	cStatistic.FRectEndDate = vEYear & "-" & TwoNumber(vEMonth) & "-" & TwoNumber(vEDay)
	cStatistic.FRectSiteName = vSiteName
	cStatistic.facademyStatistic_TimeZonelist()

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
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>시간대</td>
    <td>주문건수</td>
    <td>금액</td>
</tr>
<% if cStatistic.FTotalCount > 0 then %>
	<% For i = 0 To cStatistic.FTotalCount -1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%=cStatistic.FItemList(i).FTimeZone %>시</td>
		<td align="center"><%=cStatistic.FItemList(i).FCount %></td>
		<td align="right" style="padding-right:5px;"><%= FormatNumber(cStatistic.FItemList(i).FMaeChul,0) %></td>
		<%
		vTot_MaechulCountSum	= vTot_MaechulCountSum + CLng(FormatNumber(cStatistic.FItemList(i).FCount,0))
		vTot_MaechulPriceSum	= vTot_MaechulPriceSum + CLng(FormatNumber(cStatistic.FItemList(i).FMaeChul,0))
		%>
	</tr>
	<% Next %>
	<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
		<td bgcolor="#E6B9B8"></td>
		<td bgcolor="#E6B9B8"><%= FormatNumber(vTot_MaechulCountSum,0) %></td>
		<td align="right" style="padding-right:5px;" bgcolor="#E6B9B8"><%= FormatNumber(vTot_MaechulPriceSum,0) %></td>
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