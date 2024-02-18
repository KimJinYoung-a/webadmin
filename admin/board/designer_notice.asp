<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : [업체]공지사항
' Hieditor : 서동석 생성
'			 2023.10.23 한용민 수정(이메일발송 cdo->메일러로 변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/boardcls.asp"-->
<%
Dim ix,i, page, pgsize, TotalPage, TotalCount, prepage, nextpage, mode, nIndent, strtitle, nInstr,searchmode,search,searchString
Dim nboard, nboardFix
	search = request("search")
	searchString = request("searchString")

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

if search="" then search="title"

set nboardFix = new CBoard
	nboardFix.FTableName = "[db_board].[dbo].tbl_designer_notice"
	nboardFix.FRectFixonly = "on"
	nboardFix.FPageSize = 6
	nboardFix.design_notice

set nboard = new CBoard
	nboard.FRectFixonly= "off"

if Request("delmode") = "delete" then
	nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
	nboard.FRectIdx = Request("deletelist")
	nboard.design_notice_del
end if
if Request("SearchMode") = "search" then
	nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
	'nboard.FRectDesignerID = designer
	nboard.FPageSize = pgsize
	nboard.FCurrPage = page
	'nboard.FRectIpkumDiv4 = "on" 'ckipkumdiv4
	nboard.FRectsearch = search
	nboard.FRectsearch2 = SearchString
	nboard.design_notice
else
	'nboard.FRectItemid = itemid
	nboard.FTableName = "[db_board].[dbo].tbl_designer_notice"
	nboard.FPageSize = pgsize
	nboard.FCurrPage = page
	'nboard.FRectIpkumDiv4 = "on" 'ckipkumdiv4
	'nboard.FRectOrderSerial = orderserial
	nboard.FCurrPage = page
	nboard.design_notice
end if

%>

<script type='text/javascript'>

function urlgo(url){
	location.href=url;
}

function checkall(A,B,C){
	var X=eval("document.forms."+A+"."+B)
	for (c=0;c<X.length;c++)
	X[c].checked=C
}

function qnadelete(){
//	if (CheckMember() == true){
		var deletelist = ""
		var chk = false
		for(k=0;k<document.qnaform.elements.length;k++){
			var target=document.qnaform.elements[k]
			if(target.checked == true){
				chk = true
				deletelist = deletelist + target.value + ","
				}
		}
		if(chk == true){
			if(confirm("삭제 하시겠습니까?")){
			location = "designer_notice.asp?menupos=79&delmode=delete&page=<%=page%>&deletelist=" + deletelist;
		}
		}
		else{
			alert("먼저 삭제 목록을 선택해 주세요.");
		}
}

function gotowrite(){
	location.href="designer_notice_write.asp?idx=<% =request("idx") %>&page=<% =request("page") %>&pgsize=<% =request("pgsize") %>&menupos=79"
}

function NextPage(ipage){
	document.searchform.page.value= ipage;
	document.searchform.submit();
}

function PopEmailSend(iidx){
	var popwin = window.open('/admin/board/popemailsend.asp?id=' + iidx,'popemailsend','width=1400,height=800,scrollbars=yes,resizable=yes');
	popwin.focus();
}

</script>
</head>
<form name="searchform"  method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="SearchMode" value="search">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>"> 
<tr align="center" bgcolor="#FFFFFF" >
	<td   width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left"> 
		<select name="search" size="1"  align="absbottom" class="select">
			<option value="title">글제목</option>
			<option value="name">이름</option>
			<option value="content">내용</option>
		</select>&nbsp;
		<input name="SearchString" type="text" align="absbottom" value="<%=SearchString%>" class="text">
		<script language="javascript">searchform.search.value='<%=search%>';</script>
	</td> 
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="javascript:document.searchform.submit();">
	</td>
</tr>
</table> 
</form>
<br>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="선택항목 삭제" onclick="qnadelete();"> 
	</td>
	<td align="right">	
		<input type="button"  class="button"  value="신규등록" onclick="gotowrite();"> 
	</td>
</tr>
</table>

<form name="qnaform" method="post" style="margin:0px;">
<input type="hidden" name="page" value="1">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr>
	<td colspan="9" height="30"  bgcolor="#FFFFFF">검색결과 : 총 <font color="red"><% = nboard.FTotalCount %></font>개&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40" align="center">선택</td>
	<td width="50" align="center">번호</td>
	<td align="center">제목</td>
	<td align="center">고정글</td>
	<td align="center">팝업 공지기간</td>
	<td>카테고리</td>
	<td width="80" align="center">글쓴이</td>
	<td width="90" align="center">날짜</td>
	<td width="60" align="center">관리</td>
</tr>
<%
	if page=1 then
	for ix=0 to nboardFix.FResultCount -1
%>
  <tr align="center" bgcolor="#cccccc">
	<td><input type="checkbox" name="deletebox" value="<%= nboardFix.BoardItem(ix).FRectIdx  %>"></td>
	<td height="25"><a href="designer_notice_read.asp?idx=<%= nboardFix.BoardItem(ix).FRectIdx  %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=79"><%= nboardFix.BoardItem(ix).FRectIdx  %></a></td>
	<td><a href="designer_notice_read.asp?idx=<%= nboardFix.BoardItem(ix).FRectIdx  %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=79"><%= nboardFix.BoardItem(ix).FRectTitle %></a></td>
		<td><%= nboardFix.BoardItem(ix).Ffixnotics%></td>
		<td><%if nboardFix.BoardItem(ix).Fispopup ="Y" then%><%= nboardFix.BoardItem(ix).FpopsDate%>~<%= nboardFix.BoardItem(ix).FpopEDate%><%else%>사용안함<%end if%></td>
	<td><%= nboardFix.BoardItem(ix).Fdispcatename%></td>	
	<td><%= nboardFix.BoardItem(ix).FRectName %></td>
	<td><%= FormatDateTime(nboardFix.BoardItem(ix).Fregdate,2) %></td>
	<td>
		<input type="button" class="button" value="수정" onClick="location.href='designer_notice_modify.asp?idx=<%=nboardFix.BoardItem(ix).FRectIdx%>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=79'">
		<input type="button" class="button" value="이메일 발송" onClick="javascript:PopEmailSend('<%= nboardFix.BoardItem(ix).FRectIdx  %>')">
		
	</td>
</tr>
<%
	next
	end if
%>

<% if (nboard.FResultCount < 1) and (nboardFix.FResultCount < 1) then %>
<tr  bgcolor="#FFFFFF">
	<td colspan="12" align="center">[공지사항에 글이 없습니다.]</td>
</tr>
<% else %>
<% for ix=0 to nboard.FResultCount -1 %>
<tr bgcolor="#FFFFFF" align="center">
	<td><input type="checkbox" name="deletebox" value="<%= nboard.BoardItem(ix).FRectIdx  %>"></td>
	<td height="25"><a href="designer_notice_read.asp?idx=<%= nboard.BoardItem(ix).FRectIdx  %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=79"><%= nboard.BoardItem(ix).FRectIdx  %></a></td>
	<td><a href="designer_notice_read.asp?idx=<%= nboard.BoardItem(ix).FRectIdx  %>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=79"><%= nboard.BoardItem(ix).FRectTitle %> &nbsp;<%if nboard.BoardItem(ix).FcomCnt <>0  then%><font color="red">[<%=nboard.BoardItem(ix).FcomCnt%>]</font><%end if%></a></td>
	<td><%= nboard.BoardItem(ix).Ffixnotics%></td>
		<td><%if nboard.BoardItem(ix).Fispopup ="Y" then%><%= nboard.BoardItem(ix).FpopsDate%>~<%= nboard.BoardItem(ix).FpopEDate%><%else%>사용안함<%end if%></td>
	<td><%=nboard.BoardItem(ix).Fdispcatename%></td>
	<td><%= nboard.BoardItem(ix).FRectName %></td>
	<td><%= FormatDateTime(nboard.BoardItem(ix).Fregdate,2) %></td>
	<td>
		<input type="button" class="button" value="수정" onClick="location.href='designer_notice_modify.asp?idx=<%=nboard.BoardItem(ix).FRectIdx%>&page=<% =page %>&pgsize=<% =pgsize %>&menupos=79'">
		<input type="button" class="button" value="이메일 발송" onClick="javascript:PopEmailSend('<%= nboard.BoardItem(ix).FRectIdx  %>')">
	</td>
</tr>
<% next %>
<% end if %>

<tr  bgcolor="#FFFFFF">
<%
'response.write nboard.StartScrollPage & "<br>"
'response.write nboard.FScrollCount + nboard.StartScrollPage - 1 & "<br>"
'response.write nboard.FTotalpage
'dbget.close()	:	response.End
%>
		<td colspan="9" height="30" align="center">
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
</form>

<%
set nboardFix = Nothing
set nboard = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
