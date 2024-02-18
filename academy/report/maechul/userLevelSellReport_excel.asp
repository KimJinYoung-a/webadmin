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
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, vIsOldOrder
dim yyyymmdd1,yyymmdd2
dim fromDate,toDate
dim vSiteName, vSorting

yyyy1 = RequestCheckvar(request("yyyy1"),4)
mm1 = RequestCheckvar(request("mm1"),2)
dd1 = RequestCheckvar(request("dd1"),2)
yyyy2 = RequestCheckvar(request("yyyy2"),4)
mm2 = RequestCheckvar(request("mm2"),2)
dd2 = RequestCheckvar(request("dd2"),2)
vSiteName 	= RequestCheckvar(request("sitename"),16)
vSorting	= NullFillWith(RequestCheckvar(request("sorting"),32),"ddateD")

dim tNo, tDiv, chkOld, isBanpum
tNo = Request("tNo")
tDiv = Request("tDiv")
chkOld = Request("chkOld")
isBanpum = RequestCheckvar(Request("isBanpum"),10)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

'// 내용 접수
dim oreport
set oreport = new CUserLevelSell
	oreport.FRectSdate = fromDate
	oreport.FRectEdate = toDate
	oreport.FRectMinusInc = isBanpum
	oreport.FRectSiteName = vSiteName
	oreport.FRectSorting = vSorting
	oreport.GetLevelList

'각 비율 및 그래프 산출
dim sTotal, nTotal,  i, uTotal

if oreport.FResultCount>0 then
	for i=0 to oreport.FResultCount -1
		sTotal = sTotal + oreport.FItemList(i).FSellTotal
		nTotal = nTotal + oreport.FItemList(i).FSellCount
		uTotal = uTotal + oreport.FItemList(i).Funiqcnt
	next
end if

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>

<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>

<table width="100%" align="center" cellpadding="3" cellspacing="1" border="1" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
	<td colspan="15">
		검색결과 : <b><%= oreport.FResultCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="188" rowspan="2">회원등급</td>
	<td width="228" colspan="2">매출</td>
	<td width="228" colspan="2">건수</td>
	<td width="50" rowspan="2">Uniq고객건수</td>
	<td width="106" rowspan="2">객단가(원)</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="139">매출액(원)</td>
	<td width="89">비율(%)</td>
	<td width="139">건수</td>
	<td width="89">비율(%)</td>
</tr>
<% if oreport.FResultCount>0 then %>
	<% for i=0 to oreport.FresultCount-1 %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%=oreport.FItemList(i).GetUserLevelStr %></td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellTotal,0)%></td>
		<td><%=FormatNumber((oreport.FItemList(i).FSellTotal/sTotal)*100,2)%>%</td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellCount,0)%></td>
		<td><%=FormatNumber((oreport.FItemList(i).FSellCount/nTotal)*100,2)%>%</td>
		<td><%=FormatNumber(oreport.FItemList(i).Funiqcnt,0)%></td>
		<td><%=FormatNumber(oreport.FItemList(i).FSellAvr,0)%></td>
	</tr>
	<% next %>
	
	<tr align="center" bgcolor="#FAFAFA">
		<td>계</td>
		<td><%=FormatNumber(sTotal,0)%></td>
		<td>100%</td>
		<td><%=FormatNumber(nTotal,0)%></td>
		<td>100%</td>
		<td><%=FormatNumber(uTotal,0)%></td>
		<td><%=FormatNumber((sTotal/nTotal),0)%></td>
	</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
<%
Set oreport = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->