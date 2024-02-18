<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
<script language="JavaScript">
<!--
function NextPage(ipage){
	document.searchform.page.value= ipage;
	document.searchform.submit();
}
//-->
</script>
<%
Dim ix,i, page, pgsize
Dim TotalPage, TotalCount
Dim prepage, nextpage
Dim mode
Dim nIndent, strtitle
Dim nInstr,searchmode,search,searchString
Dim nboard
Dim nboardFix

pgsize = requestCheckVar(Request("pgsize"),10)
if pgsize="" then
	pgsize = 10
end if

page = requestCheckVar(Request("page"),10)
if page = "" then
	page = 1
end if

set nboardFix = new CBoard
nboardFix.FTableName = "[db_board].[dbo].tbl_designer_notice"
nboardFix.FRectFixonly = "on"
nboardFix.FPageSize = 7
nboardFix.FRectDesignerID = session("ssBctID")
nboardFix.design_notice_dispcate

set nboard = new CBoard
nboard.FRectFixonly = "off"

if Request("SearchMode") = "search" then
	nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
	nboard.FPageSize = pgsize
	nboard.FCurrPage = page
	nboard.FRectsearch = request("search")
	nboard.FRectsearch2 = request("SearchString")
	nboard.FRectDesignerID = session("ssBctID")
	nboard.design_notice_dispcate
else
	nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
	nboard.FPageSize = pgsize
	nboard.FCurrPage = page
	nboard.FRectDesignerID = session("ssBctID")
	nboard.design_notice_dispcate
end if


%>

<table width="780" cellspacing="0" class="a" bgcolor="#CCCCCC">
<form name="searchform"  method="get" action="">
<input type="hidden" name="page" value="1">
	<input type="hidden" name="SearchMode" value="search">
	<tr >
		<td class="a">&nbsp;
			<select name="search" size="1"  align="absbottom">
			   <option value="title">������</option>
			   <option value="name">�̸�</option>
			   <option value="content">����</option>
			</select>&nbsp;
			<input name="SearchString" type="text" align="absbottom">
		</td>
		<td class="a" align="right">
			<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0">
		</td>
	</tr>
 	</form>
</table>
<table width="780" cellspacing="1"  class="a" bgcolor=#3d3d3d>
<tr bgcolor="#FFFFFF">
	<td colspan="4" height="25" align="right">�˻���� : �� <font color="red"><% = nboard.FTotalCount %></font>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="100" align="center">��ȣ</td>
	<td align="center">����</td>
	<td width="100" align="center">�۾���</td>
	<td width="100" align="center">��¥</td>
</tr>
<form name="qnaform" method="post">
<input type="hidden" name="page" value="1">
<% for ix=0 to nboardFix.FResultCount -1 %>
<tr class="a" bgcolor="#DDDDDD">
	<td align="center" height="16">[����]</td>
	<td align="center"><a href="notics_read.asp?idx=<%= nboardFix.BoardItem(ix).FRectIdx  %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=52"><%= nboardFix.BoardItem(ix).FRectTitle %></a>
	<% if datediff("d",nboardFix.BoardItem(ix).Fregdate,now())<8 then %>
	&nbsp;<font color=red><b>new</b></font>
	<% end if %>
	</td>
	<td align="center"><%= nboardFix.BoardItem(ix).FRectName %></td>
	<td align="center"><%= FormatDateTime(nboardFix.BoardItem(ix).Fregdate,2) %></td>
</tr>
<% next %>
<% if (nboard.FResultCount < 1) and (nboardFix.FResultCount < 1) then %>
<tr bgcolor="#FFFFFF">
	<td colspan="12" align="center">[�������׿� ���� �����ϴ�.]</td>
</tr>
<% else %>
<% for ix=0 to nboard.FResultCount -1 %>
<tr class="a" bgcolor="#FFFFFF">
	<td align="center" height="22"><%= nboard.BoardItem(ix).FRectIdx  %></a></td>
	<td align="center"><a href="notics_read.asp?idx=<%= nboard.BoardItem(ix).FRectIdx  %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=52"><%= nboard.BoardItem(ix).FRectTitle %></a>
	<% if datediff("d",nboard.BoardItem(ix).Fregdate,now())<8 then %>
	&nbsp;<font color=red><b>new</b></font>
	<% end if %>
	</td>
	<td align="center"><%= nboard.BoardItem(ix).FRectName %></td>
	<td align="center"><%= FormatDateTime(nboard.BoardItem(ix).Fregdate,2) %></td>
</tr>
<% next %>
<% end if %>
</form>
	<tr bgcolor="#FFFFFF">

		<td colspan="4" height="30" align="center">
		<% if nboard.HasPreScroll then %>
			<a href="javascript:NextPage('<%= nboard.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>
		<% for ix=0 + nboard.StartScrollPage to nboard.FScrollCount + nboard.StartScrollPage - 1 %>
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
<%
set nboardFix = Nothing
set nboard = Nothing
%>
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->