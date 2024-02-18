<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research, page
dim yyyy1,mm1
dim yyyy, mm, yyyymm, yyyymm_prev, yyyymm_next
dim sellsite, searchfield, searchtext, diffType

Dim i

research = requestCheckvar(request("research"),10)
page = requestCheckvar(request("page"),10)

yyyy1   = request("yyyy1")
mm1     = request("mm1")

sellsite		= request("sellsite")
searchfield 	= request("searchfield")
searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
diffType 		= request("diffType")

if (page="") then page = 1
if (diffType="") then diffType = "DIF"


if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()) - 2)
end if

yyyymm = yyyy1 + "-" & mm1
yyyymm_prev = Left(DateSerial(yyyy1,(mm1 - 1), 1), 7)
yyyymm_next = Left(DateSerial(yyyy1,(mm1 + 1), 1), 7)


Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	oCExtJungsan.FPageSize = 25
	oCExtJungsan.FCurrPage = page

	oCExtJungsan.FRectYYYYMM = yyyymm
	oCExtJungsan.FRectDiffType = diffType

	oCExtJungsan.FRectSellSite = sellsite

	oCExtJungsan.FRectSearchField = searchfield
	oCExtJungsan.FRectSearchText = searchtext

    oCExtJungsan.GetExtJungsanDiff

	if (sellsite = "") then
		Response.write "<script>alert('제휴몰을 선택하세요.');</script>"
	end if

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popExtSiteJungsanData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("/admin/maechul/extjungsandata/popRegExtJungsanDataFile.asp","popExtSiteJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function jsExtJungsanDiffMake(sellsite, yyyymm) {
	var frm = document.frmAct;

	if (confirm("재작성하시겠습니까?") == true) {
		frm.mode.value = "extjungsandiffmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = yyyymm;

		frm.submit();
	}
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
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		&nbsp;
		제휴몰:	<% fnGetOptOutMall sellsite %>
		&nbsp;
		매출월:
		<% DrawYMBox yyyy1,mm1 %>
		&nbsp;
		조회내역:
		<input type="radio" name="diffType" value="DIF" <% if (diffType = "DIF") then %>checked<% end if %> > 오차내역
		<input type="radio" name="diffType" value="TOT" <% if (diffType = "TOT") then %>checked<% end if %> > 전체내역
		<input type="radio" name="diffType" value="SUM" <% if (diffType = "SUM") then %>checked<% end if %> > 합계내역
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		<!--
		* 검색조건 :
		<select class="select" name="searchfield">
			<option value=""></option>
			<option value="OrgOrderserial" <% if (searchfield = "OrgOrderserial") then %>selected<% end if %> >원주문번호</option>
		</select>
		<input type="text" class="text" name="searchtext" size="30" value="<%= searchtext %>">
		-->
	</td>
</tr>
</form>
</table>
<!-- 검색 끝 -->

<%

if (sellsite = "") then
	Response.write "<h5>제휴몰을 선택하세요</h5>"
end if

%>

<p>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">
		<input type="button" class="button" value="재작성(<%= sellsite %>, <%= yyyymm %>)" onClick="jsExtJungsanDiffMake('<%= sellsite %>', '<%= yyyymm %>');">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		검색결과 : <b><%= oCExtJungsan.FTotalcount %></b>
		&nbsp;
		페이지 : <b><%= page %> / <%= oCExtJungsan.FTotalPage %></b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70" rowspan="2">년월</td>
	<td width="100" rowspan="2">제휴몰</td>
	<td width="100" rowspan="2">주문번호</td>
	<td colspan="2">합계(3개월)</td>
	<td colspan="2"><%= yyyymm_prev %></td>
	<td colspan="2"><b><%= yyyymm %></b></td>
	<td colspan="2"><%= yyyymm_next %></td>
	<td rowspan="2">비고</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>SCM</td>
	<td>제휴몰</td>
	<td>SCM</td>
	<td>제휴몰</td>
	<td>SCM</td>
	<td>제휴몰</td>
	<td>SCM</td>
	<td>제휴몰</td>
</tr>
<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCExtJungsan.FItemList(i).Fyyyymm %></td>
	<td><%= oCExtJungsan.FItemList(i).GetSellSiteName %></td>
	<td><%= oCExtJungsan.FItemList(i).Forderserial %></td>

	<td width="80" align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FMeachulPriceSUM, 0) %></td>
	<td width="80" align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextMeachulPriceSUM, 0) %></td>
	<td width="80" align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FMeachulPriceSUM1, 0) %></td>
	<td width="80" align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextMeachulPriceSUM1, 0) %></td>
	<td width="80" align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FMeachulPriceSUM2, 0) %></td>
	<td width="80" align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextMeachulPriceSUM2, 0) %></td>
	<td width="80" align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FMeachulPriceSUM3, 0) %></td>
	<td width="80" align="right"><%= FormatNumber(oCExtJungsan.FItemList(i).FextMeachulPriceSUM3, 0) %></td>

	<td>

	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if oCExtJungsan.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCExtJungsan.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCExtJungsan.StartScrollPage to oCExtJungsan.FScrollCount + oCExtJungsan.StartScrollPage - 1 %>
			<% if i>oCExtJungsan.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCExtJungsan.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<form name="frmAct" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="yyyymm" value="">
</form>

<%
set oCExtJungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
