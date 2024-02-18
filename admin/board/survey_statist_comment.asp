<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<%
	Dim page, lp, div, strType, qst_sn

	qst_sn = Request("qsn")
	page = Request("page")

	'기본값 지정
	if page="" then page=1

	'// 문항 정보
	dim oSurveyQuest
	Set oSurveyQuest = new CSurvey
	oSurveyQuest.FRectSn = qst_sn
	oSurveyQuest.GetSurveyQuestCont

	'// 주관식 목록
	dim oSurvey
	Set oSurvey = new CSurvey

	oSurvey.FRectSn = qst_sn
	oSurvey.FPagesize = 15
	oSurvey.FCurrPage = page

	oSurvey.GetSurveyCommentList
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm_list.page.value=pg;
		document.frm_list.submit();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr><td><b>■ 주관식 의견 보기</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="10%" bgcolor="<%= adminColor("gray") %>">문제번호</td>
	<td width="40%" align="left"><%=qst_sn%></td>
	<td width="10%" bgcolor="<%= adminColor("gray") %>">필수여부</td>
	<td width="40%" align="left">
	<%
		if oSurveyQuest.FitemList(1).Fqst_isNull="Y" then
			Response.Write "<font color=darkblue>공란허용</font>"
		else
			Response.Write "<font color=darkred>답변필수</font>"
		end if
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">문제</td>
	<td align="left" colspan="3"><%=oSurveyQuest.FitemList(1).Fqst_content%></td>
</tr>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr><td>&nbsp;</td></tr>
</table>
<!-- 메인 목록 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_list" method="get" action="">
<input type="hidden" name="qsn" value="<%=qst_sn%>">
<input type="hidden" name="page" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		결과 : <b><%=FormatNumber(oSurvey.FTotalCount,0)%></b>
		&nbsp;
		페이지 : <b><%= page %>/<%=FormatNumber(oSurvey.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">번호</td>
	<% if oSurveyQuest.FitemList(1).Fqst_type="1" then %>
		<td width="160">선택지문</td>
		<td>내용</td>
	<% else %>
		<td width="660">내용</td>
	<% end if %>
	
</tr>
<%
	if oSurvey.FResultCount=0 then
%>
<tr>
	<td colspan="3" height="60" align="center" bgcolor="#FFFFFF">등록된 내용이 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oSurvey.FResultCount - 1
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oSurvey.FitemList(lp).Fans_sn%></td>
	<% if oSurveyQuest.FitemList(1).Fqst_type="1" then %><td><%=oSurvey.FitemList(lp).Fpoll_content%></td><% end if %>
	<td align="left"><%=oSurvey.FitemList(lp).Fans_subject%></td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<tr>
	<td colspan="3" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!-- 페이지 시작 -->
	<%
		if oSurvey.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oSurvey.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oSurvey.StartScrollPage to oSurvey.FScrollCount + oSurvey.StartScrollPage - 1

			if lp>oSurvey.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>[" & lp & "]</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
			end if

		next

		if oSurvey.HasNextScroll then
			Response.Write "&nbsp; <a href='javascript:goPage(" & lp & ")'>[next]</a>"
		else
			Response.Write "&nbsp; [next]"
		end if
	%>
	<!-- 페이지 끝 -->
	</td>
</tr>
</form>
</table>
<!-- 페이지 끝 -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->