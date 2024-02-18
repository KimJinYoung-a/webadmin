<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCompanyOpen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/company/board_cls.asp"-->
<%
	Dim brdDiv
	Dim page, SearchArea, SearchKeyword

	brdDiv = 2					'게시판 구분 (1:언론보도, 2:잡지협찬)
	page = requestCheckVar(Request("page"),9)
	searchArea = requestCheckVar(Request("searchArea"),12)
	searchKeyword = requestCheckVar(Request("searchKeyword"),32)
	if page="" then	page=1


	'// 내용 접수
	dim oBoard, lp
	Set oBoard = new CBoard

	oBoard.FPagesize = 15
	oBoard.FCurrPage = page
	oBoard.FRectBrdDiv = brdDiv
	oBoard.FRectSearchArea = SearchArea
	oBoard.FRectSearchKeyword = SearchKeyword
	
	oBoard.GetBoardList
%>
<!-- 검색 시작 -->
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="magazine_list.asp";
		document.frm.submit();
	}

	// 상세정보(수정) 페이지 이동
	function goEdit(brdsn)
	{
		document.frm.brdsn.value=brdsn;
		document.frm.page.value='<%= page %>';
		document.frm.action="magazine_edit.asp";
		document.frm.submit();
	}
//-->
</script>
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<form name="frm" method="get" action="" action="magazine_list.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="brdsn" value="">
<tr height="10" valign="bottom">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="top">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td align="right">
		<select name="SearchArea">
			<option value="">::구분::</option>
			<option value="brd_subject">제목</option>
			<option value="brd_content">내용</option>
		</select>
		<input type="text" name="SearchKeyword" size="12" value="<%=SearchKeyword%>">
		<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0" align="absmiddle">
		<script language="javascript">
			document.frm.SearchArea.value="<%=SearchArea%>";
		</script>
	</td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</form>
</table>
<!-- 검색 끝 -->
<!-- 상단 띠 시작 -->
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr><td height="1" colspan="15" bgcolor="#BABABA"></td></tr>
<tr height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td align="right">
		<table width="100%" border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
			<td>총 <%=oBoard.FtotalCount%> 개 게시물</td>
			<td align="right">page : <%= page %>/<%=oBoard.FtotalPage%></td>
		</tr>
		</table>
	</td>
		<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 상단 띠 끝 -->
<!-- 메인 목록 시작 -->
<table width="750" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#E6E6E6">
	<td width="40">번호</td>
	<td>제목</td>
	<td width="100">작성자</td>
	<td width="100">작성일</td>
	<td width="40">조회수</td>
</tr>
<%
	if oBoard.FResultCount=0 then
%>
<tr>
	<td colspan="5" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 게시물이 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oBoard.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oBoard.FitemList(lp).Fbrd_sn%></td>
	<td align="left"><a href="javascript:goEdit(<%=oBoard.FitemList(lp).Fbrd_sn%>)"><%=oBoard.FitemList(lp).Fbrd_subject%></a></td>
	<td><%=oBoard.FitemList(lp).Fuserid%></td>
	<td><%=Replace(left(oBoard.FitemList(lp).Fbrd_regdate,10),"-",".")%></td>
	<td><%=oBoard.FitemList(lp).Fbrd_hit%></td>
</tr>
<%
		next
	end if
%>
</table>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<table width="750" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr valign="bottom" height="25">
	<td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
	<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
		<tr valign="bottom">
			<td align="center">
			<!-- 페이지 시작 -->
			<%
				if oBoard.HasPreScroll then
					Response.Write "<a href='javascript:goPage(" & oBoard.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for lp=0 + oBoard.StartScrollPage to oBoard.FScrollCount + oBoard.StartScrollPage - 1

					if lp>oBoard.FTotalpage then Exit for
	
					if CStr(page)=CStr(lp) then
						Response.Write " <font color='red'>[" & lp & "]</font> "
					else
						Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
					end if

				next

				if oBoard.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
			<!-- 페이지 끝 -->
			</td>
			<td width="75" align="right"><a href="magazine_write.asp?menupos=<%=menupos%>"><img src="/images/icon_new_registration.gif" width="75" border="0"></a></td>
		</tr>
		</table>
	</td>
	<td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
	<td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_08.gif"></td>
	<td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbCompanyClose.asp" -->