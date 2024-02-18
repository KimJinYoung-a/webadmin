<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteCSOrderCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research
dim yyyy1, yyyy2, mm1, mm2, dd1, dd2
dim fromDate, toDate, tmpDate
dim sellsite

dim i, j, k, page

research = requestCheckvar(request("research"),10)

yyyy1   = requestCheckvar(request("yyyy1"), 32)
mm1     = requestCheckvar(request("mm1"), 32)
dd1     = requestCheckvar(request("dd1"), 32)
yyyy2   = requestCheckvar(request("yyyy2"), 32)
mm2     = requestCheckvar(request("mm2"), 32)
dd2     = requestCheckvar(request("dd2"), 32)

sellsite = requestCheckvar(request("sellsite"),32)

page = 1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(Day(now())) - 40)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(Day(now())) - 10)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if


Dim oCExtJungsan
set oCExtJungsan = new CxSiteCSOrder
	oCExtJungsan.FPageSize = 1000
	oCExtJungsan.FCurrPage = page

	oCExtJungsan.FRectStartDate = Left(fromDate, 10)
	oCExtJungsan.FRectEndDate = Left(toDate, 10)

	oCExtJungsan.FRectSellSite = sellsite

	if (sellsite <> "") then
		oCExtJungsan.getExtJungsanNoneList
	else
		Response.write "<script>alert('제휴몰을 선택하세요.');</script>"
	end if


%>
<!-- 검색 시작 -->
<form name="frm" method="get" action="" style="margin: 0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
	    * 쇼핑몰 :
	    <% call drawSelectBoxXSiteOrderInputPartnerCS("sellsite", sellsite) %>
		&nbsp;&nbsp;
		출고일 :
		<% DrawDateBoxdynamic yyyy1, "yyyy1", yyyy2, "yyyy2", mm1, "mm1", mm2, "mm2", dd1, "dd1", dd2, "dd2" %>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->

<p></p>

<!-- 리스트 시작 -->
<form name="frm1" method="post" style="margin: 0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="30">
		검색결과 : <b><%= oCExtJungsan.FTotalcount %></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td width="120">사이트</td>
	<td width="120">주문번호</td>
	<td width="90">상품코드</td>
	<td width="50">수량</td>
	<td width="100">입금일</td>
	<td width="200">제휴주문번호</td>
	<td width="50">제휴수량</td>
	<td>비고</td>
</tr>

<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="center" bgcolor="#FFFFFF" onmouseover='this.style.background="#F1F1F1";' onmouseout='this.style.background="#FFFFFF";' height="25">
	<td><%= oCExtJungsan.FItemList(i).FSellSite %></td>
	<td><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= oCExtJungsan.FItemList(i).Forderserial %>')"><%= oCExtJungsan.FItemList(i).Forderserial %></a></td>
	<td><%= oCExtJungsan.FItemList(i).FItemID %></td>
	<td><%= oCExtJungsan.FItemList(i).Fitemno %></td>
	<td><%= Left(oCExtJungsan.FItemList(i).Fipkumdate, 10) %></td>
	<td><%= oCExtJungsan.FItemList(i).FOutMallOrderSerial %></td>
	<td><%= oCExtJungsan.FItemList(i).FextItemno %></td>
	<td></td>
</tr>
<% next %>
</table>
</form>
<% set oCExtJungsan = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
