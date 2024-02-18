<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2012.03.22 김진영 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/artist/artist_noticeCls.asp" -->
<%
Dim i, j
Dim page,noticetype

page = request("page")
If (page = "") Then
	page = 1
End If

'==============================================================================
'공지사항
Dim oBoardNotice
Dim oNoticeFix
Dim SearchKey, SearchString, oldyn,param
Dim mm
SearchKey = Request("SearchKey")
SearchString = Request("SearchString")
oldyn = request("oldyn")
mm = request("mm")
If SearchKey="" then SearchKey="title"

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&menupos=" & menupos

Set oBoardNotice = New CBoardNotice
oBoardNotice.FPageSize = 20
oBoardNotice.FCurrPage = CInt(page)
oBoardNotice.FScrollCount = 10
oBoardNotice.FRectSearchKey = SearchKey
oBoardNotice.FRectSearchString = SearchString
oBoardNotice.List

Dim ix
%>
<script language="javascript">
function NoticeReBuild(){
	document.location.href="<%=wwwUrl%>/chtml/make_index_notice.asp?retURL=http://webadmin.10x10.co.kr/admin/board/cscenter_notice_board_list.asp?menupos=139";
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left" height="40">위치 : 
		<select onchange="location.href=this.value;" class="select">
			<option value="artist_main.asp?menupos=<%=menupos%>&mm=1">메인 멀티배너
			<option value="artist_hot_list.asp?menupos=<%=menupos%>&mm=2">HOT ARTIST
			<option value="artist_notice_board_list.asp?menupos=<%=menupos%>&mm=3" <% If mm = 3 Then response.write "selected"%> >공지사항
			<option value="artist_selectshop.asp?menupos=<%=menupos%>&mm=4">Select Shop
		</select>
	</td>
</tr>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td><img src="/images/icon_star.gif" align="absmiddle"><b>공지사항 관리</b></td>
	<td align="right"></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td colspan="2" align="right">
		<select name="SearchKey">
			<option value="title">제목</option>
			<option value="contents">내용</option>
			<option value="isusing">사용유무</option>
		</select>
		<input type="text" name="SearchString" size="30" value="<%=SearchString%>">
		<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
		</script>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 상단 띠 시작 -->
<table width="780" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr><td height="1" colspan="15" bgcolor="#BABABA"></td></tr>
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="right">
		<table width="100%" border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
			<td>총 <%=FormatNumber(oBoardNotice.FTotalCount,0)%> 개 게시물</td>
			<td align="right">page : <%= page & " / " & oBoardNotice.FTotalPage%></td>
		</tr>
		</table>
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- 상단 띠 끝 -->
<!-- 메인 목록 시작 -->
<table width="780" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#E6E6E6">
    <td width="90">번호</td>
	<td width="310">제목</td>
    <td width="160">등록일</td>
    <td width="60">고정유무</td>
     <td width="60">사용유무</td>
</tr>
<% 
IF oBoardNotice.FResultCount>0 Then 
	For i = 0 to (oBoardNotice.FResultCount - 1) 
%>
<tr bgcolor="<%= CHkIIF(oBoardNotice.results(i).Ffixyn="Y","#EFEFEF","#FFFFFF")%>" height="20">
	<td align="center"><%= oBoardNotice.results(i).Fidx %></td>
	<td><a href="artist_notice_board_modify.asp?idx=<%= oBoardNotice.results(i).Fidx & "&page=" & page & param %>"><font color="#2222FF"><%= oBoardNotice.results(i).Ftitle %></font></a></td>
	<td align="center"><%= oBoardNotice.results(i).Fregdate %></td>
	<td align="center"><%= oBoardNotice.results(i).Ffixyn %></td>
	<td align="center"><%= oBoardNotice.results(i).Fisusing %></td>
</tr>
<% 
	Next
Else
%>
<tr height="20" bgcolor="#FFFFFF">
	<td align="center" colspan="7">데이터가 없습니다.</td>
</tr>
<%
End IF
%>
</table>
<table width="780" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td align="center" class="a">
		<% if oBoardNotice.HasPreScroll then %>
			<a href="cscenter_notice_board_list.asp?page=<%= oBoardNotice.StartScrollPage-1 & param %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + oBoardNotice.StartScrollPage to oBoardNotice.FScrollCount + oBoardNotice.StartScrollPage - 1 %>
			<% if ix>oBoardNotice.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="cscenter_notice_board_list.asp?page=<%= ix & param %>">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if oBoardNotice.HasNextScroll then %>
			<a href="cscenter_notice_board_list.asp?page=<%= ix & param%>">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
	<td width="75" valign="bottom"><a href="artist_notice_board_write.asp?menupos=<%=menupos%>" onfocus="this.blur()"><img src="/images/icon_new_registration.gif" border="0"></a></td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
	<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td colspan="2" background="/images/tbl_blue_round_08.gif"></td>
	<td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->