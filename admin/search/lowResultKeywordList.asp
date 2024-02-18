<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbEVTopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/seachkeywordCls.asp" -->
<%

dim i
dim research
dim yyyy1, mm1, dd1, currDate
dim mxrectCNT, searchcnt

research 	= requestCheckvar(request("research"),32)
yyyy1 		= requestCheckvar(request("yyyy1"),32)
mm1 		= requestCheckvar(request("mm1"),32)
dd1 		= requestCheckvar(request("dd1"),32)

mxrectCNT 	= requestCheckvar(request("mxrectCNT"),32)
searchcnt 	= requestCheckvar(request("searchcnt"),32)

if (mxrectCNT = "") then
	mxrectCNT = 10
end if

if (searchcnt = "") then
	searchcnt = 10
end if

if (yyyy1 = "") then
	currDate = Now()
	yyyy1 = Year(currDate)
	mm1 = Month(currDate)
	dd1 = Day(currDate)
else
	currDate = DateSerial(yyyy1, mm1, dd1)
end if

'// ============================================================================
dim osearchKeyword

set osearchKeyword = new CSearchKeyword

osearchKeyword.FRectYYYYMMDD	= Left(currDate,10)
osearchKeyword.FRectMxrectCNT	= mxrectCNT
osearchKeyword.FRectSearchCNT	= searchcnt

osearchKeyword.GetLowResultKeywordList

%>
<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left" height="30" >
			* 날짜 : <% Call DrawOneDateBoxdynamic("yyyy1", yyyy1, "mm1", mm1, "dd1", dd1, "", "", "", "") %> ~
			&nbsp;
			* 검색횟수 : <input type="text" class="text" name="searchcnt" size="4" value="<%= searchcnt %>"> 번 이상
			&nbsp;
			* 평균검색결과수 : <input type="text" class="text" name="mxrectCNT" size="4" value="<%= mxrectCNT %>"> 개 이하
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button" value="검색" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<p />

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="21">
			검색결과 : <b><%= osearchKeyword.FResultCount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">검색어</td>
		<td width="80">검색횟수</td>
		<td width="80">평균<br />검색결과수</td>
		<td>비고</td>
	</tr>
	<%
	for i = 0 To osearchKeyword.FResultCount - 1
	%>
	<tr align="center" bgcolor="#FFFFFF">
		<td align="center" height="30">
			<%= osearchKeyword.FItemList(i).Frect %>
		</td>
		<td align="center"><%= osearchKeyword.FItemList(i).Fsumsearchcnt %></td>
		<td align="center"><%= osearchKeyword.FItemList(i).FmxrectCNT %></td>
		<td></td>
	</tr>
	<%
	next
	%>
	<% if (osearchKeyword.FResultCount = 0) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td height="30" colspan="15">
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
<!-- #include virtual="/lib/db/dbEVTclose.asp" -->
