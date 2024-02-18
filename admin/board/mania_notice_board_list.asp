<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/board/lib/classes/db2_manianewscls.asp" -->
<%

dim i, j
dim page

page = request("page")
if (page = "") then
        page = "1"
end if

'==============================================================================
'공지사항
dim boardnotice
set boardnotice = New CBoardNotice
boardnotice.PageSize = 20
boardnotice.CurrPage = CInt(page)
boardnotice.ScrollCount = 10
boardnotice.list

dim ix
%>
<script language="JavaScript">
<!--

function GoSearch(){
		document.frm.method="get";
		document.frm.submit();
}

//-->
</script>
<a href="mania_notice_board_write.asp"><font color="red">글쓰기</font></a>
<table width="760" cellspacing="1" class="a" bgcolor="#3d3d3d">
  <tr bgcolor="#DDDDFF">
    <td width="100" align="center">번호</td>
	 <td align="center">제목</td>
    <td width="160" align="center">등록일</td>
    <td width="50" align="center">사용유무</td>
  </tr>
<% for i = 0 to (boardnotice.ResultCount - 1) %>
  <tr bgcolor="#FFFFFF" height=20>
    <td width="100" align="center"><%= boardnotice.results(i).id %></td>
	<td ><a href="mania_notice_board_modify.asp?id=<%= boardnotice.results(i).id %>"><font color="#2222FF"><%= boardnotice.results(i).title %></font></a></td>
    <td width="160" align="center"><%= boardnotice.results(i).regdate %></td>
    <td width="50" align="center"><%= boardnotice.results(i).isusing %></td>
  </tr>
<% next %>
  <tr bgcolor="#FFFFFF">
    <td align="center" colspan="6">
		<% if boardnotice.HasPreScroll then %>
			<a href="?page=<%= boardnotice.StarScrollPage-1 %>">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + boardnotice.StartScrollPage to boardnotice.ScrollCount + boardnotice.StartScrollPage - 1 %>
			<% if ix>boardnotice.Totalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="?page=<%= ix %>">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if boardnotice.HasNextScroll then %>
			<a href="?page=<%= ix %>">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
  </tr>
</table>
<!-- #include virtual="/lib/db/db2close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->