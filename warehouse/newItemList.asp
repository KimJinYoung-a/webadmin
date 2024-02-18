<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  신상품리스트
' History : 서동석 생성
'			2022.02.09 한용민 수정(구매유형 디비에서 가져오게 통합작업)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/newItemCls.asp"-->
<%

dim i, j
dim purchasetype, startDT, endDT, mwdiv
dim yyyy1, yyyy2, mm1, mm2, dd1, dd2, fromDate, toDate
dim designer

yyyy1   = RequestCheckVar(request("yyyy1"),32)
mm1     = RequestCheckVar(request("mm1"),32)
dd1     = RequestCheckVar(request("dd1"),32)
yyyy2   = RequestCheckVar(request("yyyy2"),32)
mm2     = RequestCheckVar(request("mm2"),32)
dd2     = RequestCheckVar(request("dd2"),32)
designer     = RequestCheckVar(request("designer"),32)

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), CStr(Day(Now()) - 14))
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), CStr(Day(Now()) - 7))

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	yyyy2 = Cstr(Year(toDate))
	mm2 = Cstr(Month(toDate))
	dd2 = Cstr(day(toDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2)
end if

purchasetype = RequestCheckVar(request("purchasetype"), 32)
mwdiv = RequestCheckVar(request("mwdiv"), 32)

if (purchasetype = "") then
	purchasetype = "1"
	mwdiv = "M"
end if

startDT = Left(fromDate,10)
endDT = Left(toDate,10)

dim oCNewItem
set oCNewItem = new CNewItem

oCNewItem.FRectPurchaseType = purchasetype
oCNewItem.FRectStartDT = startDT
oCNewItem.FRectEndDT = endDT
oCNewItem.FRectMWDiv = mwdiv
oCNewItem.FRectMakerID = designer

oCNewItem.GetNewItemList()

dim totipgocnt, totonsellcnt, totoffsellcnt

%>
<script language='javascript'>
function popViewCurrentStock(itemgubun, itemid, itemoption) {
	var popwin;
	popwin = window.open('/admin/stock/itemcurrentstock.asp?menupos=709&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popViewCurrentStock','width=1200,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
			* 구매유형:
			<% drawPartnerCommCodeBox true,"purchasetype","purchasetype",purchasetype,"" %>
			<select class="select" name="mwdiv">
				<option value="">매입+위탁</option>
				<option value="M" <%= CHKIIF(mwdiv="M", "selected", "") %>>매입</option>
				<option value="W" <%= CHKIIF(mwdiv="W", "selected", "") %>>위탁</option>
			</select>
			* 브랜드 : <% drawSelectBoxDesignerwithName "designer", designer %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			* 기간(오픈일) :
			<% DrawDateBoxdynamic yyyy1, "yyyy1", yyyy2, "yyyy2", mm1, "mm1", mm2, "mm2", dd1, "dd1", dd2, "dd2" %>
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

* 최대 1000개까지만 표시됩니다.

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">구매유형</td>
		<td width="150">브랜드ID</td>
    	<td width="80">상품ID</td>
    	<td>상품명</td>
		<td>옵션명</td>
    	<td width="80">입고일</td>
    	<td width="80">오픈일</td>
		<td width="50">입고수</td>
		<td width="50">ON판매</td>
		<td width="50">OFF판매</td>
		<td width="50">총판매량</td>
		<td>비고</td>
    </tr>
<% for i=0 to oCNewItem.FResultCount - 1 %>
	<%
	totipgocnt = totipgocnt + oCNewItem.FItemList(i).Fipgocnt
	totonsellcnt = totonsellcnt + oCNewItem.FItemList(i).Fonsellcnt
	totoffsellcnt = totoffsellcnt + oCNewItem.FItemList(i).Foffsellcnt
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= getBrandPurchaseType(oCNewItem.FItemList(i).Fpurchasetype) %></td>
		<td><%= oCNewItem.FItemList(i).Fmakerid %></td>
		<td><%= oCNewItem.FItemList(i).Fitemid %></td>
		<td>
			<a href="javascript:popViewCurrentStock('<%= oCNewItem.FItemList(i).Fitemgubun %>', '<%= oCNewItem.FItemList(i).Fitemid %>', '<%= oCNewItem.FItemList(i).Fitemoption %>');">
				<%= oCNewItem.FItemList(i).Fitemname %>
			</a>
		</td>
		<td><%= oCNewItem.FItemList(i).Fitemoptionname %></td>
		<td><%= oCNewItem.FItemList(i).Fipgodate %></td>
		<td><%= oCNewItem.FItemList(i).FsellSTDate %></td>
		<td><%= oCNewItem.FItemList(i).Fipgocnt %></td>
		<td><%= oCNewItem.FItemList(i).Fonsellcnt %></td>
		<td><%= oCNewItem.FItemList(i).Foffsellcnt %></td>
		<td><%= (oCNewItem.FItemList(i).Fonsellcnt + oCNewItem.FItemList(i).Foffsellcnt) %></td>
		<td></td>
	</tr>
<% next %>
	<tr align="center" bgcolor="#FFFFFF">
		<td colspan="4">합계</td>
		<td><%= oCNewItem.FResultCount %></td>
		<td></td>
		<td></td>
		<td><%= totipgocnt %></td>
		<td><%= totonsellcnt %></td>
		<td><%= totoffsellcnt %></td>
		<td><%= (totonsellcnt + totoffsellcnt) %></td>
		<td></td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
