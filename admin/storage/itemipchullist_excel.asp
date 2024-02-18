<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품별입출내역 엑셀다운로드
' History : 2022.09.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/AcountItemIpChulCls.asp"-->
<%
dim gubun,designer,itemid, shopid, itemgubun, page, ipchulcode, research, arrLIst
dim IpChulMwgubun, onlineMwDiv, centermwdiv, StockMwDiv, tplgubun, purchasetype, i, sumitemno, sumSellCash, sumBuyCash, sumSuplyCash
tplgubun = request("tplgubun")
gubun       = RequestCheckVar(request("gubun"),32)
designer    = RequestCheckVar(request("designer"),32)
itemgubun   = RequestCheckVar(request("itemgubun"),2)
itemid      = RequestCheckVar(request("itemid"),9)
shopid      = RequestCheckVar(request("shopid"),32)
page        = RequestCheckVar(request("page"),10)
ipchulcode  = RequestCheckVar(request("ipchulcode"),10)
research  = RequestCheckVar(request("research"),2)
IpChulMwgubun  	= RequestCheckVar(request("IpChulMwgubun"),1)
onlineMwDiv  	= RequestCheckVar(request("onlineMwDiv"),1)
centermwdiv  	= RequestCheckVar(request("centermwdiv"),1)
StockMwDiv  	= RequestCheckVar(request("StockMwDiv"),1)
purchasetype 	= requestCheckVar(request("purchasetype"),3)
''if gubun="" then gubun="I"

if research="" and TPLGubun="" then TPLGubun="3X"
sumitemno   = 0
sumSellCash = 0
sumBuyCash  = 0
sumSuplyCash= 0

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromDate,toDate

yyyy1 = request("yyyy1")
mm1   = request("mm1")
dd1   = request("dd1")
yyyy2 = request("yyyy2")
mm2   = request("mm2")
dd2   = request("dd2")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (page="") then page=1

fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

dim oacctipchul
set oacctipchul = new CAcountItemIpChul
oacctipchul.FCurrPage = page
oacctipchul.FPageSize = 1000000
oacctipchul.FRectStartday = fromDate
oacctipchul.FRectEndday   = toDate
oacctipchul.FRectGubun   = gubun
oacctipchul.FRectDesigner = designer
oacctipchul.FRectItemGubun = itemgubun
oacctipchul.FRectItemID = itemid
oacctipchul.FRectIpChulCode = ipchulcode
oacctipchul.FtplGubun = tplgubun
oacctipchul.FRectIpChulMwgubun = IpChulMwgubun
oacctipchul.FRectOnlineMwDiv = onlineMwDiv
oacctipchul.FRectCentermwdiv = centermwdiv
oacctipchul.FRectStockMwDiv = StockMwDiv
oacctipchul.FRectBrandPurchaseType = purchasetype

if gubun<>"I" then
	oacctipchul.FRectShopid = shopid
end if

'if (designer<>"") or (itemid<>"") then
    oacctipchul.getIpChulListByItemNotPaging
'end if

if oacctipchul.FTotalCount>0 then
    arrLIst=oacctipchul.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENITEMIPCHULLIST" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
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
	<td colspan="26">
		검색결과 : <b><%= oacctipchul.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>입출코드</td>
    <td>입출구분</td>
    <td>입출고일</td>
    <% if gubun="I" then %>
    <td>업체ID</td>
    <% else %>
    <td>출고처</td>
    <% end if %>
    <td>브랜드ID</td>
    <td>상품구분</td>
    <td>상품코드</td>
    <td>옵션코드</td>
    <td>바코드</td>
    <td>상품명</td>
    <td>옵션명</td>
    <td>소비자가</td>
    <td>내역서매입가</td>
    <td>출고가</td>
    <td>수량</td>
    <td>입출시매입</td>
    <td>현재ON매입</td>
    <td>현재OF센터매입</td>
    <td>재고매입구분</td>
    <td>평균매입가(물류)</td>
    <td>평균매입가(매장)</td>
    <td>매장재고매입구분</td>
    <td>매장재고매입가</td>
    <td>구매유형</td>
    <td>카테고리</td>
</tr>
<% if oacctipchul.FResultCount>0 then %>
<% for i=0 to oacctipchul.FResultCount-1 %>
<%
    sumitemno = sumitemno + arrLIst(10,i)
    sumSellCash = sumSellCash + arrLIst(7,i)*arrLIst(10,i)
    sumBuyCash  = sumbuyCash + Null2Zero(arrLIst(9,i))*arrLIst(10,i)
    sumSuplyCash = sumSuplyCash + arrLIst(8,i)*arrLIst(10,i)
%>
<tr bgcolor="#FFFFFF">
    <td><%= arrLIst(1,i) %></td>
    <td><%= GetDivCodeName(arrLIst(2,i)) %></td>
    <td><%= arrLIst(3,i) %></td>
    <td><%= arrLIst(11,i) %></td>
    <td><%= arrLIst(15,i) %></td>
    <td><%= arrLIst(4,i) %></td>
    <td><%= arrLIst(5,i) %></td>
    <td class="txt"><%= arrLIst(6,i) %></td>
    <td class="txt"><%= arrLIst(31,i) %></td>
    <td ><%= arrLIst(12,i) %></td>
    <td><%= arrLIst(13,i) %></td>
    <td align="right"><%= FormatNumber(arrLIst(7,i),0) %></td>
    <% if arrLIst(14,i)="I" then %>
        <td align="right"><%= FormatNumber(arrLIst(8,i),0) %></td>
        <td align="right"></td>
    <% else %>
        <td align="right">
        <% if Not isNULL(arrLIst(9,i)) then %>
            <%= FormatNumber(arrLIst(9,i),0) %>
        <% end if %>
    </td>
    <td align="right"><%= FormatNumber(arrLIst(8,i),2) %></td>
    <% end if %>
    <td align="center"><%= arrLIst(10,i) %></td>
    <td align="center">
        <% if IsNULL(arrLIst(18,i)) or (arrLIst(18,i)="") or (arrLIst(18,i)=" ") then %>
        <% else %>
            <%= arrLIst(18,i) %>
        <% end if %>
    </td>
    <td align="center"><%= arrLIst(20,i) %></td>
    <td align="center"><%= arrLIst(21,i) %></td>
    <td align="center"><%= arrLIst(23,i) %></td>
    <td align="right">
        <% if Not isNULL(arrLIst(28,i)) then %>
            <%= FormatNumber(arrLIst(28,i),0) %>
        <% end if %>
    </td>
    <td align="right">
        <% if Not isNULL(arrLIst(29,i)) then %>
            <%= FormatNumber(arrLIst(29,i),0) %>
        <% end if %>
    </td>
    <td align="center"><%= arrLIst(25,i) %></td>
    <td align="right">
        <% if Not isNULL(arrLIst(26,i)) then %>
            <%= FormatNumber(arrLIst(26,i),0) %>
        <% end if %>
    </td>
    <td align="center"><%= arrLIst(22,i) %></td>
    <td align="center"><%= arrLIst(30,i) %></td>
</tr>
<%
if i mod 1000 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
%>
<tr bgcolor="#FFFFFF">
	<td colspan="11"></td>
    <td align="right"><%= FormatNumber(sumSellCash,0) %></td>
    <% if gubun="I" then %>
    <td align="right"><%= FormatNumber(sumSuplyCash,0) %></td>
    <td align="right"></td>
    <% else %>
    <td align="right"><%= FormatNumber(sumBuyCash,0) %></td>
    <td align="right"><%= FormatNumber(sumSuplyCash,2) %></td>
    <% end if %>
	<td align="center"><%= sumitemno %></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
    <td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
	<td align="center"></td>
    <td align="center"></td>
</tr>

<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="26" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>

<%
set oacctipchul = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
