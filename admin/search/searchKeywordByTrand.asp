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

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdateStr, startdateStr, nextdateStr
dim searchKeyword
dim i
dim research

research = request("research")

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

nowdateStr = CStr(now())

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Format00(2, Cstr(Month(now())))
if (dd1="") then dd1 = Format00(2, Cstr(day(now())))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Format00(2, Cstr(Month(now())))
if (dd2="") then dd2 = Format00(2, Cstr(day(now())))

if (Len(mm1) <> 2) then mm1 = Format00(2, mm1)
if (Len(dd1) <> 2) then dd1 = Format00(2, dd1)
if (Len(mm2) <> 2) then mm2 = Format00(2, mm2)
if (Len(dd2) <> 2) then dd2 = Format00(2, dd2)

startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = yyyy2 + "-" + mm2 + "-" + dd2

searchKeyword = Trim(request("searchKeyword"))

if (research = "") then
	''if (groupby = "") then groupby = "d"
end if

'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword
osearchKeyword.FRectStart 		= startdateStr
osearchKeyword.FRectEnd 		= nextdateStr
osearchKeyword.FRectKeyword		= searchKeyword

if (searchKeyword <> "") then
	osearchKeyword.getReportByTrand
end if

%>

<script language='javascript'>

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
			기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			검색어 : <input type="text" class="text" name="searchKeyword" value="<%= searchKeyword %>">
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

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100" height="30">일자</td>
		<td width="100">순위</td>
		<td>비고</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FTotalCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= osearchKeyword.FItemList(i).Fyyyymmdd %>
		</td>
		<td align="center">
			<%= osearchKeyword.FItemList(i).FkeywordRank %>
		</td>
		<td align="center">
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
