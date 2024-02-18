<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : RelateKeywordLink_List.asp
' Discription : 카테고리 관련 키워드 목록
' History : 2008.03.28 허진원 생성
'			2022.07.05 한용민 수정(isms취약점조치, 표준코딩으로변경)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/CategoryCls.asp"-->
<%
	Dim page, SearchKey, SearchString

	page = Request("page")
	SearchKey = Request("SearchKey")
	SearchString = Request("SearchString")
	if page="" then	page=1


	'// 내용 접수
	dim oRelate, lp
	Set oRelate = new CRelateList

	oRelate.FPagesize = 15
	oRelate.FCurrPage = page
	oRelate.FRectCDL = request("cdl")
	oRelate.FRectCDM = request("cdm")
	oRelate.FRectCDS = request("cds")
	oRelate.FRectsearchKey = searchKey
	oRelate.FRectsearchString = searchString
	
	oRelate.GetRelateLinkList
%>
<!-- 검색 시작 -->
<script type='text/javascript'>
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.action="RelateKeywordLink_list.asp";
		document.frm.submit();
	}

	// 아이템 상세정보(수정) 페이지 이동
	function goEdit(rid)
	{
		document.frm.rid.value=rid;
		document.frm.page.value='<%= page %>';
		document.frm.action="RelateKeywordLink_Edit.asp";
		document.frm.submit();
	}

	// 아이템 삭제 실행
	function goDel(rid)
	{
		if(confirm("[" + rid + "]번 관련키워드를 삭제하시겠습니까?\n\n※ 완료시 완전히 삭제되며 복구할 수 없습니다.")) {
			document.frm.rid.value=rid;
			document.frm.mode.value="delete";
			document.frm.action="DoRelate_Process.asp";
			document.frm.submit();
		}
	}

	// 신규등록 페이지로 이동
	function goAddItem()  {
		self.location="RelateKeywordLink_Edit.asp?menupos=<%=menupos%>";
	}
//-->
</script>
<form name="frm" method="get" action="" action="RelateKeywordLink_list.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="rid" value="">
<input type="hidden" name="mode" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td>
			<table width="100%" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td><!-- #include virtual="/common/module/categoryselectbox.asp"--></td>
				<td align="right">
					키워드 :
					<select class="select" name="SearchKey">
						<option value="">::구분::</option>
						<option value="linkCode">링크코드</option>
						<option value="linkKeyword">키워드</option>
						<option value="linkURL">링크</option>
					</select>
					<input type="text" class="text" name="SearchString" size="20" value="<%=SearchString%>">
					<script language="javascript">
						document.frm.SearchKey.value="<%=SearchKey%>";
					</script>
				</td>
			</tr>
			</table>
		</td>
		<td width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="submit" class="button_s" value="검색">
		</td>
	</tr>
</table>
</form>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding:10 0 10 0;">
<tr>
	<td align="right">
		<input type="button" class="button" value="신규등록" onClick="goAddItem()">
	</td>
</tr>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		검색결과 : <b><%=oRelate.FtotalCount%></b>
		&nbsp;
		페이지 : <b><%= page %> / <%=oRelate.FtotalPage%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>코드</td>
	<td>카테고리</td>
	<td>키워드</td>
	<td>링크</td>
	<td>수정/삭제</td>
</tr>
<%
	if oRelate.FResultCount=0 then
%>
<tr>
	<td colspan="5" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 아이템이 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oRelate.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oRelate.FitemList(lp).Flinkcode%></td>
	<td align="left">
	<%
		Response.Write oRelate.FitemList(lp).FCDL_nm
		if Not(oRelate.FitemList(lp).FCDM_nm="" or isNull(oRelate.FitemList(lp).FCDM_nm)) then Response.Write " > " & oRelate.FitemList(lp).FCDM_nm
		if Not(oRelate.FitemList(lp).FCDS_nm="" or isNull(oRelate.FitemList(lp).FCDS_nm)) then Response.Write " > " & oRelate.FitemList(lp).FCDS_nm
	%>
	</td>
	<td align="left"><%= ReplaceBracket(oRelate.FitemList(lp).FlinkKeyword) %></td>
	<td align="left"><%= ReplaceBracket(oRelate.FitemList(lp).FlinkURL) %></td>
	<td>
		<input type="button" value="수정" class="button" onClick="goEdit(<%=oRelate.FitemList(lp).Flinkcode%>)">
		<input type="button" value="삭제" class="button" onClick="goDel(<%=oRelate.FitemList(lp).Flinkcode%>)">
	</td>
</tr>	
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15" align="center">
	<!-- 페이지 시작 -->
	<%
		if oRelate.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oRelate.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oRelate.StartScrollPage to oRelate.FScrollCount + oRelate.StartScrollPage - 1

			if lp>oRelate.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>" & lp & "</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>" & lp & "</a> "
			end if

		next

		if oRelate.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->