<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 설문관리
' Hieditor : 허진원 생성
'			 2022.07.08 한용민 수정(isms취약점보안조치, 표준코드로변경)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<%
	Dim page, lp, div, using, strDiv
	page = requestCheckVar(getNumeric(request("page")),10)
	div = Request("div")
	using = requestCheckVar(request("using"),1)

	'기본값 지정
	if page="" then page=1
	if using="" then using="Y"


	'// 설문 목록
	dim oSurvey
	Set oSurvey = new CSurvey

	oSurvey.FPagesize = 15
	oSurvey.FCurrPage = page
	oSurvey.FRectUsing = using
	oSurvey.FRectDiv = div
	
	oSurvey.GetSurveyList
%>
<script type='text/javascript'>
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}
//-->
</script>
<!-- 상단 검색폼 시작 -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">검색조건</td>
	<td align="left">
		구분 <select name="div" class="select">
			<option value="">전체</option>
			<option value="1">업체</option>
			<option value="2">직원</option>
		</select>
		/ 상태 <select name="using" class="select">
			<option value="N">삭제</option>
			<option value="Y">사용</option>
		</select>
		<script language="javascript">
			document.frm.div.value="<%=div%>";
			document.frm.using.value="<%=using%>";
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="검색">
	</td>
</tr>
</table>
</form>
<br>
<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>&nbsp;</td>
	<td align="right" style="padding:4 0 4 0"><input type="button" class="button" value="신규등록" onClick="window.open('survey_write.asp','SurveyPop','width=1400,height=768')"></td>
</tr>
</table>
<!-- 액션 끝 -->
<!-- 메인 목록 시작 -->
<form name="frm_list" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="div" value="<%=div%>">
<input type="hidden" name="using" value="<%=using%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		검색결과 : <b><%=FormatNumber(oSurvey.FTotalCount,0)%></b>
		&nbsp;
		페이지 : <b><%= page %>/<%=FormatNumber(oSurvey.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>일련번호</td>
	<td>설문제목</td>
	<td>구분</td>
	<td>시작일</td>
	<td>종료일</td>
	<td>등록일</td>
	<td>상태</td>
</tr>
<%
	if oSurvey.FResultCount=0 then
%>
<tr>
	<td colspan="7" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 설문이 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oSurvey.FResultCount - 1
			'구분
			Select Case oSurvey.FitemList(lp).Fsrv_div
				Case "1"
					strDiv = "업체"
				Case "2"
					strDiv = "직원"
			end Select
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oSurvey.FitemList(lp).Fsrv_sn%></td>
	<td>
		<a href="survey_qst_list.asp?sn=<%=oSurvey.FitemList(lp).Fsrv_sn%>&menupos=<%=menupos%>">
		<%= ReplaceBracket(oSurvey.FitemList(lp).Fsrv_subject) %></a>
	</td>
	<td><%=strDiv%></td>
	<td><%=left(oSurvey.FitemList(lp).Fsrv_startDt,10)%></td>
	<td><%=left(oSurvey.FitemList(lp).Fsrv_endDt,10)%></td>
	<td><%=left(oSurvey.FitemList(lp).Fsrv_regdate,10)%></td>
	<td><%= oSurvey.FitemList(lp).getSurveyState %></td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<tr>
	<td colspan="7" align="center" bgcolor="<%= adminColor("tabletop") %>">
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
</table>
</form>
<!-- 페이지 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->