<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<%
	Dim page, lp, div, strType, srv_sn

	srv_sn = Request("sn")
	page = Request("page")
	div = Request("div")

	'기본값 지정
	if page="" then page=1

	'// 설문내용 접수
	dim oSurveyMaster
	Set oSurveyMaster = new CSurvey

	oSurveyMaster.FRectSn = srv_sn
	
	oSurveyMaster.GetSurveyStatistCont

	'// 설문문항 목록
	dim oSurveyQuestion
	Set oSurveyQuestion = new CSurvey

	oSurveyQuestion.FRectSn = srv_sn
	oSurveyQuestion.FPagesize = 8
	oSurveyQuestion.FCurrPage = page
	oSurveyQuestion.FRectOrder = "asc"

	oSurveyQuestion.GetSurveyQstStatist
%>
<script language="javascript">
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm_list.page.value=pg;
		document.frm_list.submit();
	}

	// 기타의견,주관식 답변 팝업
	function popCommentView(qstSn)
	{
		window.open("survey_statist_comment.asp?qsn="+qstSn,"popComView","width=720,height=600,scrollbars=yes");
	}
//-->
</script>
<script language="javascript" src="/lib/util/chart/FusionCharts.js"></script>
<!-- 설문정보 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="10%" bgcolor="<%= adminColor("gray") %>">설문번호</td>
	<td width="40%" align="left"><%=srv_sn%></td>
	<td width="10%" bgcolor="<%= adminColor("gray") %>">상태</td>
	<td width="40%" align="left">
	<%
		if oSurveyMaster.FitemList(1).Fsrv_isusing="Y" then
			if date()<oSurveyMaster.FitemList(1).Fsrv_startDt then
				Response.Write "<font color=darkgreen>대기</font>"
			elseif date()>oSurveyMaster.FitemList(1).Fsrv_endDt then
				Response.Write "<font color=darkorange>종료</font>"
			else
				Response.Write "<font color=darkblue>진행중</font>"
			end if
		else
			Response.Write "<font color=darkred>삭제</font>"
		end if
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">기간</td>
	<td align="left"><%=left(oSurveyMaster.FitemList(1).Fsrv_startDt,10) & " ~ " & left(oSurveyMaster.FitemList(1).Fsrv_endDt,10)%></td>
	<td bgcolor="<%= adminColor("gray") %>">참여수</td>
	<td align="left"><%=oSurveyMaster.FitemList(1).FansCnt%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">제목</td>
	<td align="left" colspan="3"><%=oSurveyMaster.FitemList(1).Fsrv_subject%></td>
</tr>
</table>
<!-- 설문정보 끝 -->
<!-- 설문액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>&nbsp;</td>
</tr>
</table>
<!-- 설문액션 끝 -->
<!-- 메인 목록 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_list" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sn" value="<%=srv_sn%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="div" value="<%=div%>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		검색결과 : <b><%=FormatNumber(oSurveyQuestion.FTotalCount,0)%></b>
		&nbsp;
		페이지 : <b><%= page %>/<%=FormatNumber(oSurveyQuestion.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">번호</td>
	<td width="50">형태</td>
	<td>문항</td>
	<td width="410">답변통계</td>
	<td width="60">기타</td>
</tr>
<%
	if oSurveyQuestion.FResultCount=0 then
%>
<tr>
	<td colspan="5" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 문항이 없습니다.</td>
</tr>
<%
	else
		for lp=0 to oSurveyQuestion.FResultCount - 1
			'구분
			Select Case oSurveyQuestion.FitemList(lp).Fqst_type
				Case "1"
					strType = "객관식"
				Case "2"
					strType = "주관식"
				Case "3"
					strType = "단답형"
				Case "9"
					strType = "구분자"
			end Select
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oSurveyQuestion.FitemList(lp).Fqst_sn%></td>
	<td><%=strType%></td>
	<td align="left"><%=oSurveyQuestion.FitemList(lp).Fqst_content%></td>
	<% if strType="객관식" then %>
	<td>
		<div id="chartdiv<%=lp%>" align="center"></div>
		<script type="text/javascript">	
			var chart = new FusionCharts("/lib/util/chart/MSBar2D.swf", "chartdiv<%=lp%>", "400", "150", "0", "0");
			chart.setDataURL("survey_answer_xml.asp?qsn=<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>");
			chart.render("chartdiv<%=lp%>");
		</script>
	</td>
	<td><a href="javascript:popCommentView(<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>)">[의견보기]</a></td>
	<% elseif strType="주관식" then %>
	<td colspan="2"><a href="javascript:popCommentView(<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>)">[주관식 답변 보기]</a></td>
	<% elseif strType="단답형" then %>
	<td colspan="2"><a href="javascript:popCommentView(<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>)">[단답형 답변 보기]</a></td>
	<% end if %>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<tr>
	<td colspan="5" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!-- 페이지 시작 -->
	<%
		if oSurveyQuestion.HasPreScroll then
			Response.Write "<a href='javascript:goPage(" & oSurveyQuestion.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
		else
			Response.Write "[pre] &nbsp;"
		end if

		for lp=0 + oSurveyQuestion.StartScrollPage to oSurveyQuestion.FScrollCount + oSurveyQuestion.StartScrollPage - 1

			if lp>oSurveyQuestion.FTotalpage then Exit for

			if CStr(page)=CStr(lp) then
				Response.Write " <font color='red'>[" & lp & "]</font> "
			else
				Response.Write " <a href='javascript:goPage(" & lp & ")'>[" & lp & "]</a> "
			end if

		next

		if oSurveyQuestion.HasNextScroll then
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
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->