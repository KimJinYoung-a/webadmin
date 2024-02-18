<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 상품별주문리스트
' History : 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/stock/ordersheetcls.asp"-->
<%
dim page, research, yyyy1,mm1,dd1,yyyy2,mm2,dd2, fromDate,toDate, i, baljucode, itemid, makerid, mwdiv
dim purchasetype, blinkcode, datetype, statecd, tplgubun, productidx, arrLIst
dim sumRealItemnoSellcash, sumRealItemnoBuycash, sumBaljuItemno, sumRealItemno, sumCheckItemno
    page = RequestCheckVar(getNumeric(trim(request("page"))),10)
    research = RequestCheckVar(trim(request("research")),2)
    yyyy1 = RequestCheckVar(trim(request("yyyy1")),4)
    mm1   = RequestCheckVar(trim(request("mm1")),2)
    dd1   = RequestCheckVar(trim(request("dd1")),2)
    yyyy2 = RequestCheckVar(trim(request("yyyy2")),4)
    mm2   = RequestCheckVar(trim(request("mm2")),2)
    dd2   = RequestCheckVar(trim(request("dd2")),2)
    baljucode   = RequestCheckVar(trim(request("baljucode")),32)
    itemid      = requestCheckvar(trim(request("itemid")),1500)
    makerid   = RequestCheckVar(trim(request("makerid")),32)
    mwdiv       = requestCheckvar(trim(request("mwdiv")),10)
    purchasetype = RequestCheckVar(getNumeric(trim(request("purchasetype"))),10)
    blinkcode   = RequestCheckVar(trim(request("blinkcode")),32)
    datetype   = RequestCheckVar(trim(request("datetype")),32)
    statecd   = RequestCheckVar(trim(request("statecd")),10)
    tplgubun   = RequestCheckVar(trim(request("tplgubun")),32)
    productidx = RequestCheckVar(getNumeric(trim(request("productidx"))),10)

if datetype="" or isnull(datetype) then datetype="regdate"
if (yyyy1="") then yyyy1 = Cstr(Year(dateadd("d",-7,date())))
if (mm1="") then mm1 = Cstr(Month(dateadd("d",-7,date())))
if (dd1="") then dd1 = Cstr(day(dateadd("d",-7,date())))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (page="") then page=1
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if itemid<>"" then
	dim iA ,arrTemp,arrItemid
  itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp)
		if Trim(arrTemp(iA))<>"" and isNumeric(Trim(arrTemp(iA))) then
			arrItemid = arrItemid & Trim(arrTemp(iA)) & ","
		end if
		iA = iA + 1
	loop

	if len(arrItemid)>0 then
		itemid = left(arrItemid,len(arrItemid)-1)
	else
		if Not(isNumeric(itemid)) then
			itemid = ""
		end if
	end if
end if

dim oItemOrder
set oItemOrder = new COrderSheet
oItemOrder.FCurrPage = page
oItemOrder.FPageSize = 1000000
oItemOrder.FRectStartDate = fromDate
oItemOrder.FRectEndDate   = toDate
oItemOrder.FRectbaljucode   = baljucode
oItemOrder.FRectblinkcode   = blinkcode
oItemOrder.FRectItemid       = itemid
oItemOrder.FRectmakerid       = makerid
oItemOrder.FRectmwdiv       = mwdiv
oItemOrder.FRectBrandPurchaseType = purchasetype
oItemOrder.FRectdatetype = datetype
oItemOrder.FRectstatecd = statecd
oItemOrder.FRecttplgubun = tplgubun
oItemOrder.FRectproductidx = productidx
oItemOrder.GetItemOrderListNotPaging

if oItemOrder.FTotalCount>0 then
    arrLIst=oItemOrder.fArrLIst
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENITEMORDERLIST" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
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
	<td colspan="27">
		검색결과 : <b><%= oItemOrder.FTotalCount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>주문코드</td>
    <td>원가IDX</td>
    <td>품의번호</td>
    <td>주문일</td>
    <td>입고요청일</td>
    <td>주문상태</td>
    <td>관련입고코드</td>
    <td>브랜드ID</td>
    <td>상품구분</td>
    <td>상품코드</td>
    <td>옵션코드</td>
    <td>바코드</td>
    <td>범용바코드</td>
    <td>업체관리코드</td>
    <td>상품명</td>
    <td>옵션명</td>
    <td>소비자가</td>
    <td>확정총소비자가</td>
    <td>내역서매입가</td>
    <td>확정총매입가</td>
    <td>매입구분</td>
    <td>주문수량</td>
    <td>확정수량</td>
    <td>검품수량</td>
    <td>구매유형</td>
    <td>카테고리</td>
    <td>최종입고월(물류)</td>
</tr>
<% if isarray(arrLIst) then %>
<%
sumRealItemnoSellcash=0
sumRealItemnoBuycash=0
sumBaljuItemno=0
sumRealItemno=0
sumCheckItemno=0
for i=0 to ubound(arrLIst,2)
'sumRealItemnoSellcash = sumRealItemnoSellcash + (arrLIst(16,i)*arrLIst(20,i))
'sumRealItemnoBuycash = sumRealItemnoBuycash + (arrLIst(17,i)*arrLIst(20,i))
'sumBaljuItemno = sumBaljuItemno + arrLIst(19,i)
'sumRealItemno = sumRealItemno + arrLIst(20,i)
'sumCheckItemno = sumCheckItemno + arrLIst(21,i)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td><%= arrLIst(0,i) %></td>
    <td><%= arrLIst(1,i) %></td>
    <td><%= arrLIst(2,i) %></td>
    <td><%= left(arrLIst(3,i),10) %></td>
    <td><%= arrLIst(4,i) %></td>
    <td><%= arrLIst(5,i) %></td>
    <td><%= arrLIst(6,i) %></td>
    <td><%= arrLIst(7,i) %></td>
    <td><%= arrLIst(8,i) %></td>
    <td class="txt"><%= arrLIst(9,i) %></td>
    <td class="txt"><%= arrLIst(10,i) %></td>
    <td class="txt"><%= arrLIst(11,i) %></td>
    <td class="txt"><%= arrLIst(12,i) %></td>
    <td class="txt"><%= arrLIst(13,i) %></td>
    <td align="left"><%= arrLIst(14,i) %></td>
    <td align="left"><%= arrLIst(15,i) %></td>
    <td align="right"><%= FormatNumber(arrLIst(16,i),0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(16,i)*arrLIst(20,i),0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(17,i),0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(17,i)*arrLIst(20,i),0) %></td>
    <td><%= mwdivName(arrLIst(18,i)) %></td>
    <td align="right"><%= FormatNumber(arrLIst(19,i),0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(20,i),0) %></td>
    <td align="right"><%= FormatNumber(arrLIst(21,i),0) %></td>
    <td><%= arrLIst(22,i) %></td>
    <td><%= arrLIst(23,i) %></td>
    <td class="txt"><%= arrLIst(24,i) %></td>
</tr>
<%
if i mod 300 = 0 then
    Response.Flush		' 버퍼리플래쉬
end if
next
%>
<!--<tr bgcolor="#FFFFFF" align="center">
    <td colspan=17>합계</td>
    <td align="right"><%'= FormatNumber(sumRealItemnoSellcash,0) %></td>
    <td></td>
    <td align="right"><%'= FormatNumber(sumRealItemnoBuycash,0) %></td>
    <td></td>
    <td align="right"><%'= FormatNumber(sumBaljuItemno,0) %></td>
    <td align="right"><%'= FormatNumber(sumRealItemno,0) %></td>
    <td align="right"><%'= FormatNumber(sumCheckItemno,0) %></td>
    <td colspan=3></td>
</tr>-->
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="27" align="center">[검색결과가 없습니다.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
set oItemOrder = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->