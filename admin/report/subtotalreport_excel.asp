<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 구매금액별건수 엑셀다운로드
' Hieditor : 2019.09.11 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2, yyyymmdd1,yyymmdd2, fromDate,toDate, i,p1,p2,pro,pro2, totcnt, totsum
dim ck_joinmall,ck_ipjummall,research, ck_tendeliverExists, oldlist, pricegubunMin, pricegubunMax, pricegubun
	yyyy1 = RequestCheckVar(request("yyyy1"),4)
	mm1 = RequestCheckVar(request("mm1"),2)
	dd1 = RequestCheckVar(request("dd1"),2)
	yyyy2 = RequestCheckVar(request("yyyy2"),4)
	mm2 = RequestCheckVar(request("mm2"),2)
	dd2 = RequestCheckVar(request("dd2"),2)
	research = RequestCheckVar(request("research"),2)
	ck_joinmall = RequestCheckVar(request("ck_joinmall"),2)
	ck_ipjummall = RequestCheckVar(request("ck_ipjummall"),2)
	ck_tendeliverExists = RequestCheckVar(request("ck_tendeliverExists"),2)
	oldlist = RequestCheckVar(request("oldlist"),2)

if research<>"on" then
	if ck_joinmall="" then ck_joinmall="on"
	if ck_ipjummall="" then ck_ipjummall="on"
end if

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = "1"

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)

dim oreport
set oreport = new CJumunMaster
	oreport.FRectFromDate = fromDate
	oreport.FRectToDate = toDate
	oreport.FRectJoinMallNotInclude = ck_joinmall
	oreport.FRectExtMallNotInclude = ck_ipjummall
	oreport.FRectOldJumun = oldlist
	oreport.FRectTenDeliverExists = ck_tendeliverExists
	oreport.SearchMallSellrePort6

Response.Buffer=true
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<table width="100%" align="center" border=1>
<tr align="center">
	<td>최소금액</td>
	<td>최대금액</td>
	<td>주문건수</td>
    <td>비중</td>
    <td>매출</td>
    <td>비중</td>
</tr>
<% if oreport.FResultCount > 0 then %>
    <% for i=0 to oreport.FResultCount-1 %>
    <%
        pro = 0
        if oreport.maxc<>0 then
            p1 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.maxt*100)
            p2 = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.maxc*100)
            if oreport.FTotalsellcnt<>0 then
                pro = Clng(oreport.FMasterItemList(i).Fsellcnt/oreport.FTotalsellcnt*100)
            end if

            if oreport.Ftotalmoney<>0 then
                pro2 = Clng(oreport.FMasterItemList(i).Fselltotal/oreport.Ftotalmoney*100)
            end if
        end if
        totcnt = totcnt + oreport.FMasterItemList(i).Fsellcnt
        totsum = totsum + oreport.FMasterItemList(i).Fselltotal

        pricegubunMin=""
        pricegubunMax=""
        pricegubun = trim(oreport.FMasterItemList(i).Fsitename)
        if pricegubun<>"" then
            pricegubun = split(pricegubun,"~")
            if isarray(pricegubun) then
                pricegubunMin = pricegubun(0)
                pricegubunMax = pricegubun(1)
            end if
        end if
        if i mod 5000 = 0 then
            Response.Flush		' 버퍼리플래쉬
        end if
    %>
    <tr bgcolor="#FFFFFF">
        <td align="left"><%= pricegubunMin %></td>
        <td align="left"><%= pricegubunMax %></td>
        <td><%= FormatNumber(oreport.FMasterItemList(i).Fsellcnt,0) %></td>
        <td><%= pro %>%</td>
        <td><%= FormatNumber(oreport.FMasterItemList(i).Fselltotal,0) %></td>
        <td><%= pro2 %>%</td>
    </tr>
    <% next %>
    <tr bgcolor="#FFFFFF" align="center">
        <td colspan="6" align="right">
            총건수 : <%= FormatNumber(totcnt,0) %>
            총금액 : <%= FormatNumber(totsum,0) %>
            객단가 :
            <% if totcnt<>0 then %>
            <%= FormatNumber(CLng(totsum/totcnt),0) %>
            <% end if %>
        </td>
    </tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="6" align="center" class="page_link">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>

<%
set oreport = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
