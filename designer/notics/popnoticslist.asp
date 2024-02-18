<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->

<script language="JavaScript">

function PopNotice(v){
    var popwin = window.open("/designer/notics/notics_read.asp?idx=" + v ,"PopNotice","width=650,height=600,scrollbars=yes,resizable=yes");
	popwin.focus();
}


<!--

function NextPage(ipage){
	document.searchform.page.value= ipage;
	document.searchform.submit();
}

//-->
</script>
<%

response.expires = 0



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
'nboard.FRectDesignerID = designer
nboard.FPageSize = pgsize
nboard.FCurrPage = page
'nboard.FRectIpkumDiv4 = "on" 'ckipkumdiv4
nboard.FRectsearch = request("search")
nboard.FRectsearch2 = request("SearchString")
nboard.FRectDesignerID = session("ssBctID")
nboard.design_notice_dispcate
else
nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
'nboard.FRectDesignerID = designer
nboard.FPageSize = pgsize
nboard.FCurrPage = page
'nboard.FRectIpkumDiv4 = "on" 'ckipkumdiv4
'nboard.FRectOrderSerial = orderserial
nboard.FCurrPage = page
nboard.FRectDesignerID = session("ssBctID")
nboard.design_notice_dispcate
end if


%>



<p>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<form name="searchform"  method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="SearchMode" value="search">
   	<tr height="10" valign="bottom">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top">
	        	<select name="search" size="1"  align="absbottom">
			   		<option value="title">글제목</option>
			   		<option value="name">이름</option>
			   		<option value="content">내용</option>
				</select>
				&nbsp
				<input name="SearchString" type="text" align="absbottom">
				&nbsp
				검색결과 : 총 <font color="red"><% = nboard.FTotalCount %></font>개
		    </td>
	        <td valign="top" align="right">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->


<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    	<td width="100">번호</td>
    	<td>제목</td>
      	<td width="100">작성자</td>
      	<td width="100">작성일</td>
    </tr>

	<form name="qnaform" method="post">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">

	<% for ix=0 to nboardFix.FResultCount -1 %>
	<tr class="a" bgcolor="<%= adminColor("pink") %>">
		<td align="center" height="16">[공지]</td>
		<td align="center"><a href="javascript:PopNotice('<%= nboardFix.BoardItem(ix).FRectIdx  %>');"><%= nboardFix.BoardItem(ix).FRectTitle %></a>
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
		<td colspan="12" align="center">[공지사항에 글이 없습니다.]</td>
	</tr>
	<% else %>
	<% for ix=0 to nboard.FResultCount -1 %>
	<tr class="a" bgcolor="#FFFFFF">
		<td align="center" height="22"><%= nboard.BoardItem(ix).FRectIdx  %></a></td>
		<td align="center">
		    <a href="javascript:PopNotice(<%= nboard.BoardItem(ix).FRectIdx %>);"><%= nboard.BoardItem(ix).FRectTitle %></a>
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
</table>


<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="<%= adminColor("topbar") %>" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
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
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->



<%
set nboardFix = Nothing
set nboard = Nothing
%>

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->