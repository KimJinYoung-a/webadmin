<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*10		' 10분
%>
<%
'###########################################################
' Description : 재고월령 엑셀다운로드
' History : 이상구 생성
'			 2023.10.11 한용민 수정(csv파일 -> 엑셀파일 생성으로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls_2.asp"-->
<%
Dim reqYYYYMM, reqStrplace, reqsysorreal, reqbPriceGbn, reqmygubun, reqYYYY, IsUsingV2, strNoType, strPriceType, strYearMonth
Dim AdmPath, appPath, sNow, sY, sM, sD, sH, sMi, sS, sDateName, FileName, fso, tFile, FTotCnt, FTotPage, FCurrPage, sqlStr
dim i, ArrRows, headLine, ojaego, arrLIst, tmpPrice
	reqYYYYMM = RequestCheckVar(request("exYYYY"),4)&"-"&RequestCheckVar(request("exMM"),4)
	reqStrplace = RequestCheckVar(request("stplace"),1)
	reqsysorreal = RequestCheckVar(request("sysorreal"),10)
	reqbPriceGbn = RequestCheckVar(request("bPriceGbn"),1)
	reqmygubun = RequestCheckVar(request("mygubun"),1)
	reqYYYY = RequestCheckVar(request("exYYYY"),4)
	IsUsingV2 = RequestCheckVar(request("v2"),10)

if (IsUsingV2 = "") then
	IsUsingV2 = "Y"
end if

set ojaego = new CMonthlyMaeipLedge
ojaego.FCurrPage = 1
ojaego.FPageSize = 1000000
ojaego.frectreqYYYYMM = reqYYYYMM
ojaego.frectreqStrplace = reqStrplace
ojaego.frectIsUsingV2 = IsUsingV2
ojaego.GetJeagoOverValueListNotPaging

if ojaego.FTotalCount>0 then
    arrLIst=ojaego.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENMonthlyStockList" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '버퍼사용여부
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="FFFFFF">
	<td colspan="40">
		검색결과 : <b><%= ojaego.FTotalCount %></b>
	</td>
</tr>
<%
strNoType		= "실사(+불량)"
strPriceType	= "작성시매입가"
strYearMonth	= "1-3개월,4개월~6개월,7개월~12개월,1년~2년,2년초과"

if (reqsysorreal = "sys") then
    strNoType = "시스템"
end if
if (reqbPriceGbn = "V") then
    strPriceType = "평균매입가"
end if
%>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>부서</td>
    <td>구매유형</td>
    <td>매입구분</td>
    <td>브랜드</td>
    <td>매장</td>
    <td>구분</td>
    <td>상품코드</td>
    <td>옵션코드</td>
    <td>바코드</td>
    <td>상품명</td>
    <td>옵션명</td>
    <td>최종입고일</td>
    <td>수량(시스템)</td>
    <td>공급가(<%= strPriceType %>)</td>

    <% if (reqmygubun = "Y") then %>
        <td><%= reqYYYY %></td>
        <td><%= reqYYYY - 1 %></td>
        <td><%= reqYYYY - 2 %></td>
        <td><%= reqYYYY - 3 %></td>
    <% else %>
        <td>1-3개월</td>
        <td>4개월~6개월</td>
        <td>7개월~12개월</td>
        <td>1년~2년</td>
        <td>2년초과</td>
    <% end if %>

    <td>NULL</td>
    <td>합계</td>
    <td>전시카테고리</td>
    <td>전시카테고리</td>
    <td>관리카테고리</td>
    <td>관리카테고리</td>
    <td>소비자가</td>
    <td>현재판매가</td>
    <td>현재판매여부</td>
    <td>과세구분</td>
    <td>현재센터매입구분</td>
    <td>현재매입구분</td>
</tr>
<% if isarray(arrLIst) then %>
<%
for i=0 to ubound(arrLIst,2)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td><%= arrList(1,i) %></td>
    <td><%= arrList(2,i) %></td>
    <td><%= arrList(3,i) %></td>
    <td class="txt"><%= arrList(4,i) %></td><% ' 브랜드 %>
    <td><%= trim(arrList(12,i)) %></td>
    <td><%= arrList(6,i) %></td>
    <td><%= arrList(7,i) %></td>
    <td class="txt"><%= arrList(8,i) %></td><% ' 옵션코드 %>
    <td class="txt"><%= arrList(40,i) %></td><% ' 바코드 %>
    <td><%= arrList(9,i) %></td>
    <td><%= arrList(10,i) %></td>
    <td class="txt"><%= arrList(11,i) %></td><% ' 최종입고일 %>
    <td><%= arrList(13,i) %></td>

    <td>
        <% if (reqbPriceGbn = "V") then %>
            <%= arrList(16,i) %>
            <% tmpPrice = arrList(16,i) %>
        <% else %>
            <%= arrList(15,i) %>
            <% tmpPrice = arrList(15,i) %>
        <% end if %>
    </td>
    <% if (reqsysorreal = "sys") then %>
        <% if (reqmygubun = "Y") then %>
            <td><%= arrList(22,i)*tmpPrice %></td>
            <td><%= arrList(23,i)*tmpPrice %></td>
            <td><%= arrList(24,i)*tmpPrice %></td>
            <td><%= arrList(25,i)*tmpPrice %></td>
        <% else %>
            <td><%= arrList(17,i)*tmpPrice %></td>
            <td><%= arrList(18,i)*tmpPrice %></td>
            <td><%= arrList(19,i)*tmpPrice %></td>
            <td><%= arrList(20,i)*tmpPrice %></td>
            <td><%= arrList(21,i)*tmpPrice %></td>
        <% end if %>

        <td><%= arrList(26,i)*tmpPrice %></td>
        <td><%= arrList(13,i)*tmpPrice %></td>
        <td><%= arrList(38,i) %></td>
    <% else %>
        <% if (reqmygubun = "Y") then %>
            <td><%= arrList(22+10,i)*tmpPrice %></td>
            <td><%= arrList(23+10,i)*tmpPrice %></td>
            <td><%= arrList(24+10,i)*tmpPrice %></td>
            <td><%= arrList(25+10,i)*tmpPrice %></td>
        <% else %>
            <td><%= arrList(17+10,i)*tmpPrice %></td>
            <td><%= arrList(18+10,i)*tmpPrice %></td>
            <td><%= arrList(19+10,i)*tmpPrice %></td>
            <td><%= arrList(20+10,i)*tmpPrice %></td>
            <td><%= arrList(21+10,i)*tmpPrice %></td>
        <% end if %>

        <td><%= arrList(26+10,i)*tmpPrice %></td>
        <td><%= arrList(13+1,i)*tmpPrice %></td>
        <td><%= arrList(38,i) %></td>
    <% end if %>
    
    <td><%= arrList(41,i) %></td>

    <% ' 관리카테고리 %>
    <td><%= arrList(42,i) %></td>
    <td><%= arrList(43,i) %></td>
    <td><%= arrList(44,i) %></td>
    <td><%= arrList(45,i) %></td>
    <td><%= arrList(46,i) %></td>
    <td><%= arrList(47,i) %></td>
    <td><%= arrList(48,i) %></td>
    <td><%= arrList(49,i) %></td>
</tr>
<%
if i mod 300 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
%>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="40" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set ojaego = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->