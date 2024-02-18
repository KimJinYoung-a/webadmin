<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim yyyy1,mm1,dd1
dim yyyymmdd1
dim nowdateStr, startdateStr, nextdateStr
dim i
dim research

research = request("research")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")

nowdateStr = CStr(now())

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Format00(2, Cstr(Month(now())))
if (dd1="") then dd1 = Format00(2, Cstr(day(now())))

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1


if (research = "") then
	''if (groupby = "") then groupby = "d"
end if

'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword
osearchKeyword.FRectBaseDate	= startdateStr

osearchKeyword.getReportByPopularAndDay

%>

<script language='javascript'>

function popOpenTrand(yyyy1, yyyy2, mm1, mm2, dd1, dd2, currKeyword) {
	if ((yyyy1 == yyyy2) && (mm1 == mm2) && (dd1 == dd2)) {
		var startDate = new Date(yyyy1, (mm1 - 1), (dd1 - 7));
		yyyy1 = startDate.getFullYear();
		mm1 = startDate.getMonth() + 1;
		dd1 = startDate.getDate();
	}

	var popwin = window.open("/admin/search/searchKeywordByTrand.asp?yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&dd1=" + dd1 + "&yyyy2=" + yyyy2 + "&mm2=" + mm2 + "&dd2=" + dd2 + "&searchKeyword=" + currKeyword,"popOpenTrand","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

function popOpenRelated(yyyy1, yyyy2, mm1, mm2, dd1, dd2, currKeyword) {
	if ((yyyy1 == yyyy2) && (mm1 == mm2) && (dd1 == dd2)) {
		var startDate = new Date(yyyy1, (mm1 - 1), (dd1 - 7));
		yyyy1 = startDate.getFullYear();
		mm1 = startDate.getMonth() + 1;
		dd1 = startDate.getDate();
	}

	var popwin = window.open("/admin/search/searchKeywordByRelated.asp?yyyy1=" + yyyy1 + "&mm1=" + mm1 + "&dd1=" + dd1 + "&yyyy2=" + yyyy2 + "&mm2=" + mm2 + "&dd2=" + dd2 + "&searchKeyword=" + currKeyword,"popOpenRelated","width=800 height=600 scrollbars=yes resizable=yes");
	popwin.focus();
}

</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" height="30" >
			일자 : <% DrawOneDateBox yyyy1,mm1,dd1 %>
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left" height="30">
			(1시간 지연 데이터)
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p>

* 일주일은 금요일 00:00 부터 목요일 23:59 까지 입니다.

<p>

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="50" height="30">순위</td>
		<td width="100">전일대비</td>
		<td width="100">전주대비</td>
		<td>검색어</td>
		<td width="100">검색추세</td>
		<td width="100">연관검색어</td>
		<td width="100">검색횟수</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FTotalCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= (i + 1) %>
		</td>
		<td align="center">
			<% if abs(osearchKeyword.FItemList(i).FrankPrevDay - (i + 1)) >= 10 then%><b><% end if %>
			<% if ((osearchKeyword.FItemList(i).FrankPrevDay - (i + 1)) > 0) then %>
				<font color="blue">+
			<% elseif ((osearchKeyword.FItemList(i).FrankPrevDay - (i + 1)) < 0) then %>
				<font color="red">-
			<% end if %>
			<%= abs(osearchKeyword.FItemList(i).FrankPrevDay - (i + 1)) %>
		</td>
		<td align="center">
			<% if abs(osearchKeyword.FItemList(i).FrankPrevWeek - (i + 1)) >= 10 then%><b><% end if %>
			<% if ((osearchKeyword.FItemList(i).FrankPrevWeek - (i + 1)) > 0) then %>
				<font color="blue">+
			<% elseif ((osearchKeyword.FItemList(i).FrankPrevWeek - (i + 1)) < 0) then %>
				<font color="red">-
			<% end if %>
			<%= abs(osearchKeyword.FItemList(i).FrankPrevWeek - (i + 1)) %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).FcurrKeyword %>
		</td>
		<td align="center">
			<a href="javascript:popOpenTrand('<%= yyyy1 %>', '<%= yyyy1 %>', '<%= mm1 %>', '<%= mm1 %>', '<%= dd1 %>', '<%= dd1 %>', '<%= osearchKeyword.FItemList(i).FcurrKeyword %>')">보기</a>
		</td>
		<td align="center">
			<a href="javascript:popOpenRelated('<%= yyyy1 %>', '<%= yyyy1 %>', '<%= mm1 %>', '<%= mm1 %>', '<%= dd1 %>', '<%= dd1 %>', '<%= osearchKeyword.FItemList(i).FcurrKeyword %>')">보기</a>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).Fcount %>
		</td>
	</tr>
	<%
	next
	%>
	<% if (osearchKeyword.FTotalCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="3">
			검색결과가 없습니다.
		</td>
	</tr>
	<% end if %>
</table>
<%
set osearchKeyword = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
