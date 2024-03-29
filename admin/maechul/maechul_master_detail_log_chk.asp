<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 제휴몰 클래스
' Hieditor : 2011.04.22 이상구 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/lib/classes/maechul/maechulLogCls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research, page
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim yyyy, mm, dd
dim fromDate ,toDate, tmpDate
dim targetGbn

Dim i

research = requestCheckvar(request("research"),10)
page = requestCheckvar(request("page"),10)

yyyy1   = request("yyyy1")
mm1     = request("mm1")
dd1     = request("dd1")
yyyy2   = request("yyyy2")
mm2     = request("mm2")
dd2     = request("dd2")

targetGbn		= requestCheckvar(request("targetGbn"),10)

if (page="") then page = 1
if (targetGbn = "") then
	targetGbn = "ON"
end if

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 7)

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

Dim oCMaechulPaymentLog
set oCMaechulPaymentLog = new CMaechulLog
	oCMaechulPaymentLog.FPageSize = 100
	oCMaechulPaymentLog.FCurrPage = page

	oCMaechulPaymentLog.FRectStartdate = fromDate
	oCMaechulPaymentLog.FRectEndDate = toDate

	oCMaechulPaymentLog.GetMaechulMasterDetailLogCheck

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function jsSetDate(yyyy, mm, dd) {
	var frm = document.frm;

	frm.yyyy1.value = yyyy;
	frm.mm1.value = mm;
	frm.dd1.value = dd;

	frm.yyyy2.value = yyyy;
	frm.mm2.value = mm;
	frm.dd2.value = dd;

    document.frm.submit();
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
        <!--
		&nbsp;
		매출구분 : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "" %>
        -->
		&nbsp;
		결제일자(처리일자) :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
	</td>
	<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<p>

	* 처리일자를 하루로 제한하면 주문번호가 표시됩니다.<br />

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oCMaechulPaymentLog.FTotalcount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCMaechulPaymentLog.FTotalPage %></b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="60">구분</td>
	<td width="80">결제일<br>(처리일)</td>
	<td width="100">주문번호</td>
	<td width="100">마스터 결제액</td>
	<td width="80">디테일 결제액</td>
	<td width="80">오차</td>

	<td>비고</td>
</tr>

<% for i=0 to oCMaechulPaymentLog.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCMaechulPaymentLog.FItemList(i).FtargetGbn %></td>
	<td>
		<a href="javascript:jsSetDate('<%= Left(oCMaechulPaymentLog.FItemList(i).Factdate, 4) %>', '<%= Right(Left(oCMaechulPaymentLog.FItemList(i).Factdate, 7), 2) %>', '<%= Right(oCMaechulPaymentLog.FItemList(i).Factdate, 2) %>')">
			<%= oCMaechulPaymentLog.FItemList(i).Factdate %>
		</a>
	</td>
	<td><%= oCMaechulPaymentLog.FItemList(i).Forderserial %></td>
	<td align="right"><%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FtotalOrderMaechulPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FtotalpayreqPrice, 0) %></td>
	<td align="right"><%= FormatNumber(oCMaechulPaymentLog.FItemList(i).FtotalOrderMaechulPrice - oCMaechulPaymentLog.FItemList(i).FtotalpayreqPrice, 0) %></td>

	<td>

	</td>
</tr>
<% next %>

</form>
</table>

<%
set oCMaechulPaymentLog = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
