<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2008.04.29 한용민 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/board/boardnoticecls.asp" -->
<%

dim i, j
dim page,noticetype

page = request("page")
if (page = "") then
	page = 1
end if

noticetype = request("noticetype")
'==============================================================================
'공지사항
dim oBoardNotice
dim oNoticeFix

dim SearchKey, SearchString, oldyn,param

SearchKey = Request("SearchKey")
SearchString = Request("SearchString")
oldyn = request("oldyn")

if SearchKey="" then SearchKey="title"

param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&oldyn="& oldyn &"&noticetype=" & noticetype & "&menupos=" & menupos

'set oNoticeFix = new CBoardNotice
'oNoticeFix.FRectFixonly = "on"
'oNoticeFix.FPageSize = 10
'oNoticeFix.FScrollCount = 10
'oNoticeFix.FCurrPage = 1
'if page=1 and SearchString="" and noticetype="" and oldyn<>"" then
'	oNoticeFix.list
'end if

set oBoardNotice = New CBoardNotice

oBoardNotice.FPageSize = 20
oBoardNotice.FCurrPage = CInt(page)
oBoardNotice.FScrollCount = 10
oBoardNotice.FRectnoticetype = noticetype

if oldyn<>"" Then
	oBoardNotice.FRectOldYn = oldyn
	oBoardNotice.FRectSearchKey = SearchKey
	oBoardNotice.FRectSearchString = SearchString
	oBoardNotice.FRectNoticeOrder = "1"
	oBoardNotice.List
Else
	'oBoardNotice.FRectFixonly = ""
	oBoardNotice.FRectNoticeOrder = "7"
	oBoardNotice.getNoticsList
End if



dim ix
%>
<script language="javascript">
function NoticeReBuild(){
	document.location.href="<%=wwwUrl%>/chtml/make_index_notice.asp?retURL=http://webadmin.10x10.co.kr/admin/board/cscenter_notice_board_list.asp?menupos=139";
}
</script>
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
	<td>
		<img src="/images/icon_star.gif" align="absmiddle">
		<b>공지사항 관리</b>
	</td>
	<td align="right"><img src="http://webadmin.10x10.co.kr/images/button_reload.gif" alt="메인공지사항 새로고침" onclick="NoticeReBuild()" style="cursor:pointer"></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td colspan="2" align="right">
		<input type="checkbox" name="oldyn" value="on" <% IF oldyn<>"" Then response.write "checked" %>>전체보기 /
		구분
		<select name="noticetype">
			<option value="">선택</option>
			<!--<option value="01">전체공지</option> 2015리뉴얼에서 빠짐. 이상준대리.//-->
			<option value="02">안내</option>
			<option value="03">이벤트공지</option>
			<option value="04">배송공지</option>
			<option value="05">당첨자공지</option>
			<option value="06">CultureStation</option>
		</select>
		/ 키워드
		<select name="SearchKey">
			<option value="title">제목</option>
			<option value="contents">내용</option>
		</select>
		<input type="text" name="SearchString" size="12" value="<%=SearchString%>">
		<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
		<script language="javascript">
			document.frm.SearchKey.value="<%=SearchKey%>";
			document.frm.noticetype.value="<%=noticetype%>";
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
    <td width="90">공지유형</td>
	<td width="310">제목</td>
    <td width="80">유효시작일</td>
    <td width="80">유효종료일</td>
    <td width="160">등록일</td>
    <td width="60">고정유무</td>
</tr>
<%
if false then 'oNoticeFix.FResultCount = 0 then
	for i = 0 to (oNoticeFix.FResultCount - 1)
%>
  <tr bgcolor="#EFEFEF" height=20>
    <td align="center"><%= oNoticeFix.results(i).NoticeTypeName %></td>
	<td ><a href="cscenter_notice_board_modify.asp?id=<%= oNoticeFix.results(i).Fid & "&page=" & page & param %>"><font color="#2222FF"><%= oNoticeFix.results(i).Ftitle %></font></a></td>
    <td align="center"><%= oNoticeFix.results(i).Fyuhyostart %></td>
    <td align="center"><%= oNoticeFix.results(i).Fyuhyoend %></td>
    <td align="center"><%= oNoticeFix.results(i).Fregdate %></td>
    <td align="center"><%= oNoticeFix.results(i).Ffixyn %></td>
  </tr>
<%
	next
end if
%>
<% 
IF oBoardNotice.FResultCount>0 Then 
	For i = 0 to (oBoardNotice.FResultCount - 1) 
%>
  <tr bgcolor="<%= CHkIIF(oBoardNotice.results(i).Ffixyn="Y","#EFEFEF","#FFFFFF")%>" height="20">
    <td align="center"><%= oBoardNotice.results(i).NoticeTypeName %></td>
	<td ><a href="cscenter_notice_board_modify.asp?id=<%= oBoardNotice.results(i).Fid & "&page=" & page & param %>"><font color="#2222FF"><%= oBoardNotice.results(i).Ftitle %></font></a></td>
    <td align="center"><%= oBoardNotice.results(i).Fyuhyostart %></td>
    <td align="center"><%= oBoardNotice.results(i).Fyuhyoend %></td>
    <td align="center"><%= oBoardNotice.results(i).Fregdate %></td>
    <td align="center"><%= oBoardNotice.results(i).Ffixyn %></td>
  </tr>
<% 
	Next 
End IF%>
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
	<td width="75" valign="bottom"><a href="cscenter_notice_board_write.asp" onfocus="this.blur()"><img src="/images/icon_new_registration.gif" border="0"></a></td>
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