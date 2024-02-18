<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : 현재재고다운로드
' History	:  이상구 생성
'              2021.12.16 한용민 수정
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/datelib.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls.asp"-->
<%
dim yyyy1,mm1,stockPlace,i,stockPlacename, day1after, arrlist,noti, shopid
    yyyy1       = requestCheckvar(request("yyyy1"),10)
    mm1       	= requestCheckvar(request("mm1"),10)
    stockPlace  = requestCheckvar(request("stockPlace"),10)
    shopid  	= requestCheckvar(request("shopid"),32)
    day1after       = requestCheckvar(request("day1after"),2)
noti=""
'if yyyy1="" or mm1="" or stockPlace="" then
if stockPlace="" then
    response.write "<script type='text/javascript'>"
    response.write "    alert('재고위치 구분값이 없습니다.');"
    response.write "</script>"
    dbget.close() : response.end
end if
if stockPlace="L" then
    stockPlacename = "물류"
elseif stockPlace="S" then
    stockPlacename = "매장"
else
    response.write "<script type='text/javascript'>"
    response.write "    alert('지정된 구분값이 아닙니다.');"
    response.write "</script>"
    dbget.close() : response.end
end if
dim ostock
set ostock = new CMonthlyMaeipLedge
ostock.FRectYYYYMM = yyyy1 & "-" & mm1
ostock.FRectStockPlace = stockPlace
ostock.FRectShopID = shopid
ostock.FPageSize = 1
ostock.FCurrPage = 100000

if stockPlace="L" then
    ''ostock.Getmonthlystock_notpaging
    ostock.GetCurrentStockList

' 태섭 2013-01 테스트 가능
elseif stockPlace="S" then
    ''ostock.Getmonthlyshopstock_notpaging
    ostock.GetCurrentShopstockList
end if
arrlist = ostock.farrlist

Response.Buffer=true
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN_재고실사_"& stockPlacename &noti&"_" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 12">
<style type="text/css">
 td {font-size:8.0pt;}
 .txt {mso-number-format:"\@";}
 .num {mso-number-format:"0";}
 .prc {mso-number-format:"\#\,\#\#0";}
</style>
</head>
<body>
<!--[if !excel]>　　<![endif]-->
<div align=center x:publishsource="Excel">

<table width="100%" border="1" align="center" class="a csH15" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="table-layout:fixed">
    <tr bgcolor="<%= adminColor("tabletop") %>" align="center">
        <% if stockPlace="L" then %>
            <td>재고속성</td>
            <td>랙코드앞자리</td>
            <td>랙코드</td>
            <td>보조랙코드</td>
            <td>브랜드</td>
            <td>구분</td>
            <td>상품코드</td>
            <td>옵션코드</td>
            <td>바코드</td>
            <td>상품명</td>
            <td>옵션명</td>
            <td>최종입고일(물류)</td>
            <td>시스템재고</td>
            <td>시스템재고(BLK)</td>
            <td>시스템재고(AGV)</td>
            <td>평균매입가(부가세제외)</td>
            <td>합계</td>
            <td>누적오차</td>
            <td>누적불량</td>
            <td>실사유효재고</td>
            <td>실사유효재고(BLK)</td>
            <td>실사유효재고(AGV)</td>
            <td>1일이후변동</td>
            <td>1일이후오차</td>
            <td>예상수량</td>
            <td>실사여부</td>
            <td>비고</td>
        <% else %>
            <td>매장</td>
            <td>브랜드</td>
            <td>구분</td>
            <td>상품코드</td>
            <td>옵션코드</td>
            <td>상품명</td>
            <td>옵션명</td>
            <td>최종입고일</td>
            <td>수량</td>
            <td>공급가</td>
            <td>합계</td>
            <td>바코드</td>
            <td>실사오차</td>
            <td>이동중수량</td>
            <td>반품중수량</td>
            <td>예상수량</td>
            <td>실사여부</td>
        <% end if %>
    </tr>
<% if isarray(arrlist) then %>
<% for i = 0 to ubound(arrlist,2) %>
    <tr bgcolor="#FFFFFF" align="center" >
        <% if stockPlace="L" then %>
            <td class="txt"><%= arrlist(0,i) %></td>
            <td><%= arrlist(1,i) %></td>
            <td><%= arrlist(2,i) %></td>
            <td class="txt"><%= arrlist(3,i) %></td>
            <td align="left"><%= arrlist(4,i) %></td>
            <td><%= arrlist(5,i) %></td>
            <td><%= arrlist(6,i) %></td>
            <td class="txt"><%= arrlist(7,i) %></td>
            <td class="txt"><%= arrlist(8,i) %></td>
            <td align="left"><%= arrlist(9,i) %></td>
            <td align="left"><%= arrlist(10,i) %></td>
            <td class="txt"><%= arrlist(11,i) %></td>
            <td><%= arrlist(12,i) %></td>
            <td><%= arrlist(13,i) %></td>
            <td><%= arrlist(14,i) %></td>
            <td><%= arrlist(15,i) %></td>
            <td class="txt"><%= arrlist(16,i) %></td>
            <td><%= arrlist(17,i) %></td>
            <td><%= arrlist(18,i) %></td>
            <td><%= arrlist(19,i) %></td>
            <td><%= arrlist(20,i) %></td>
            <td><%= arrlist(21,i) %></td>
            <td><%= arrlist(22,i) %></td>
            <td><%= arrlist(23,i) %></td>
            <td></td>
            <td></td>
            <td></td>
        <% else %>
            <td><%= arrlist(16,i) %></td>
            <td class="txt"><%= arrlist(0,i) %></td>
            <td><%= arrlist(1,i) %></td>
            <td><%= arrlist(2,i) %></td>
            <td class="txt"><%= arrlist(3,i) %></td>
            <td align="left"><%= arrlist(4,i) %></td>
            <td align="left"><%= arrlist(5,i) %></td>
            <td class="txt"><%= arrlist(6,i) %></td>
            <td><%= arrlist(7,i) %></td>
            <td><%= arrlist(8,i) %></td>
            <td><%= arrlist(9,i) %></td>
            <td class="txt"><%= arrlist(10,i) %></td>
            <td><%= arrlist(11,i) %></td>
            <td><%= arrlist(12,i) %></td>
            <td><%= arrlist(13,i) %></td>
            <td><%= arrlist(14,i) %></td>
            <td>
                <% if day1after="" then %>
                    <%= arrlist(15,i) %>
                <% else %>
                <% end if %>
            </td>
        <% end if %>
    </tr>
<%
if i mod 1000 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
end if
%>

</table>
</div>
</body>
</html>
<%
set ostock = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
