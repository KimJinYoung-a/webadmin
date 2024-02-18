<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/boardcls.asp"-->
<%

response.expires = 0

	Dim ix,i, page, pgsize
	Dim TotalPage, TotalCount
	Dim prepage, nextpage
	Dim mode
	Dim nIndent, strtitle
	Dim nInstr,searchmode,search,searchString
    Dim nboard

	if Request("pgsize")="" then
		pgsize = 10
	else
		pgsize = Request("pgsize")
	end if

	if Request("page") = "" then
		page = 1
	else
		page = Request("page")
	end if


set nboard = new CBoard

if Request("delmode") = "delete" then
	nboard.FTableName = "[db_board].[10x10].tbl_designer_board"
	nboard.FRectIdx = Request("deletelist")
	nboard.design_board_del
elseif Request("SearchMode") = "search" then
	nboard.FTableName = "[db_board].[10x10].tbl_designer_board"
	nboard.FRectDesignerID = session("ssBctID")
	nboard.FPageSize = pgsize
	nboard.FCurrPage = page
	'nboard.FRectIpkumDiv4 = "on" 'ckipkumdiv4
	nboard.FRectsearch = request("search")
	nboard.FRectsearch2 = request("SearchString")
	nboard.admin_design_board
else
	nboard.FTableName = "[db_board].[10x10].tbl_designer_board"
	nboard.FRectDesignerID = session("ssBctID")
	nboard.FPageSize = pgsize
	nboard.FCurrPage = page
	'nboard.FRectIpkumDiv4 = "on" 'ckipkumdiv4
	'nboard.FRectOrderSerial = orderserial
	nboard.FCurrPage = page
	nboard.admin_design_board
end if

%>

<script language="JavaScript">
<!--

function gotowrite(){
location.href="designer_board_write.asp?page=<% =request("page") %>&pgsize=<% =request("pgsize") %>&menupos=90"
}

function NextPage(ipage){
	document.searchform.page.value= ipage;
	document.searchform.submit();
}

//-->
</script>
</head>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
<form name="searchform"  method="get" action="">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="SearchMode" value="search">
	<tr>
		<td class="a">&nbsp;
			<select name="search" size="1"  align="absbottom">
			   <option value="title">글제목</option>
			   <option value="name">이름</option>
			   <option value="content">내용</option>
			</select>&nbsp;
			<input name="SearchString" type="text" align="absbottom">
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
 	</form>
</table>
<table width="100%" border="1" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td colspan="4" height="25" align="right">검색결과 : 총 <font color="red"><% = nboard.FTotalCount %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr >
	<td width="100" align="center">번호</td>
	<td align="center">제목</td>
	<td width="100" align="center">글쓴이</td>
	<td width="100" align="center">날짜</td>
</tr>
<% if nboard.FResultCount < 1 then %>
<tr>
	<td colspan="12" align="center">[게시판에 글이 없습니다.]</td>
</tr>
<% else %>
<form name="qnaform" method="post">
<% for ix=0 to nboard.FResultCount -1 %>
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<tr class="a">
	<td align="center" height="25"><%= nboard.BoardItem(ix).FRectIdx  %></a></td>
	<td>&nbsp;
				<% if nboard.BoardItem(ix).FRectLevel > 0 then
						if nboard.BoardItem(ix).FRectLevel > 1 then
						 nIndent = nboard.BoardItem(ix).FRectLevel -1
						end if %>
				  <font color="blue">&nbsp;[reply]</font>
				<%end if%>
<% if nboard.BoardItem(ix).FRectDeleteyn = "Y" then %>
<STRIKE><%= db2html(nboard.BoardItem(ix).FRectTitle) %></STRIKE>
<% else %>
	<a href="designer_board_read.asp?idx=<%= nboard.BoardItem(ix).FRectIdx  %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=90"><%= nboard.BoardItem(ix).FRectTitle %></a>
<% end if %>
	</td>
	<td align="center"><%= db2html(nboard.BoardItem(ix).FRectName) %></td>
	<td align="center"><%= FormatDateTime(nboard.BoardItem(ix).Fregdate,2) %></td>
</tr>
<% next %>
<% end if %>
</form>
	<tr>
		<td colspan="4" height="30" align="center">
		<% if nboard.HasPreScroll then %>
			<a href="javascript:NextPage('<%= nboard.StarScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + nboard.StarScrollPage to nboard.FScrollCount + nboard.StarScrollPage - 1 %>
			<% if ix>nboard.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(ix) then %>
			<font color="red">[<%= ix %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= ix %>')">[<%= ix %>]</a>
			<% end if %>
		<% next %>

		<% if nboard.HasNextScroll then %>
			<a href="javascript:NextPage('<%= ix %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
		</td>
	</tr>
</table>
<table>
<tr>
	<td><input type="button" value="게시판 쓰기" onclick="gotowrite();"></td>
</tr>
</table>
<%

set nboard = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->