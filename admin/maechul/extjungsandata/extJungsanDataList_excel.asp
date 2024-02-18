<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->

<%	
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="+"ext_JungsanDataList_excel"+".xls"


dim research, page
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim yyyy, mm, dd
dim fromDate ,toDate, tmpDate
dim sellsite, jungsantype, searchfield, searchtext

Dim i

research = requestCheckvar(request("research"),10)
page = requestCheckvar(request("page"),10)

yyyy1   = request("yyyy1")
mm1     = request("mm1")
dd1     = request("dd1")
yyyy2   = request("yyyy2")
mm2     = request("mm2")
dd2     = request("dd2")

sellsite		= requestCheckvar(request("sellsite"),32)
jungsantype		= requestCheckvar(request("jungsantype"),32)
searchfield 	= requestCheckvar(request("searchfield"),32)
searchtext 		= Replace(Replace(requestCheckvar(request("searchtext"),32), "'", ""), Chr(34), "")

dim extjdate : extjdate = requestCheckvar(request("extjdate"),8) ''YYYYMMDD
dim mimap : mimap = requestCheckvar(request("mimap"),10) 
dim vatyn : vatyn = requestCheckvar(request("vatyn"),1) 
dim retonly : retonly = requestCheckvar(request("retonly"),10) 
dim errexists : errexists = requestCheckvar(request("errexists"),10) 
dim dotview : dotview = requestCheckvar(request("dotview"),10) 
dim FormatDotNo : FormatDotNo = 0
dim exceptcost0 : exceptcost0 = requestCheckvar(request("exceptcost0"),10) 
if (dotview<>"") then FormatDotNo = 2

if (extjdate="") then 
    extjdate = replace(LEFT(dateAdd("d",-1,now()),10),"-","")
end if

if (page="") then page = 1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, DateAdd("m",1,toDate))
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if

Dim arrRows
Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	''oCExtJungsan.FPageSize = 25
	''oCExtJungsan.FCurrPage = page

	oCExtJungsan.FRectStartdate = fromDate
	oCExtJungsan.FRectEndDate = toDate

	oCExtJungsan.FRectSellSite = sellsite
	oCExtJungsan.FRectJungsanType = jungsantype

	oCExtJungsan.FRectSearchField = searchfield
	oCExtJungsan.FRectSearchText = searchtext

	oCExtJungsan.FRectMimap = mimap
	oCExtJungsan.FRectVatYn = vatyn
	oCExtJungsan.FRectReturnOnly = retonly
	oCExtJungsan.FRectErrexists = errexists
	oCExtJungsan.FRectExceptItemCostZero = exceptcost0
    arrRows = oCExtJungsan.GetExtJungsanExcelDown

	set oCExtJungsan = Nothing
%>

<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="100">제휴몰</td>
	<td width="100">매출일자</td>
	<td width="150">제휴<br>주문번호</td>
	<td width="100">제휴<br>주문순번</td>
	<td width="150">제휴<br>원주문번호</td>
	<td width="40">수량</td>
	<td width="100">판매가</td>
	<td width="100">제휴부담쿠폰</td>
	<td width="100">텐텐부담쿠폰</td>
	<td width="100">쿠폰가</td>
	<td width="100"><b>매출금액</b></td>
	<td width="100">수수료</td>
	<td width="100">정산금액</td>
	<td width="150">원주문번호</td>
	<td width="100">상품코드</td>
	<td width="60">옵션코드</td>
</tr>

<% if isArray(ArrRows) then %>
<% For i =0 To UBound(ArrRows,2) %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= ArrRows(0,i) %></td>
	<td><%= ArrRows(1,i) %></td>
	<td style="mso-number-format:\@"><%= ArrRows(2,i) %></td>
	<td><%= ArrRows(3,i) %></td>
	<td><%= ArrRows(4,i) %></td>
	<td><%= ArrRows(5,i) %></td>
	<td align="right"><%= FormatNumber(ArrRows(6,i), 0) %></td>
	<td align="right"><%= FormatNumber(ArrRows(7,i), 0) %></td>
	<td align="right"><%= FormatNumber(ArrRows(8,i), 0) %></td>
	<td align="right"><%= FormatNumber(ArrRows(9,i), 0) %></td>
	<td align="right"><b><%= FormatNumber(ArrRows(10,i), 0) %></b></td>
	<td align="right"><%= FormatNumber(ArrRows(11,i), 0) %></td>
	<td align="right"><%= FormatNumber(ArrRows(12,i), 0) %></td>
	<td style="mso-number-format:\@"><%= ArrRows(13,i) %></td>
	<td><%= ArrRows(14,i) %></td>
	<td style="mso-number-format:\@"><%= ArrRows(15,i) %></td>
</tr>
<% if (i mod 1000 = 0) then response.flush %>
<% next %>
</table>
<% end if %>

<!-- #include virtual="/lib/db/dbclose.asp" -->
