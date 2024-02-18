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
	Dim page, lp, div, using, strType, strDel, srv_sn
	srv_sn = Request("sn")
	page = requestCheckVar(getNumeric(request("page")),10)
	div = Request("div")
	using = requestCheckVar(request("using"),1)

	'기본값 지정
	if page="" then page=1
	if using="" then using="Y"

	'// 설문내용 접수
	dim oSurveyMaster
	Set oSurveyMaster = new CSurvey
	oSurveyMaster.FRectSn = srv_sn
	oSurveyMaster.GetSurveyCont

	'// 설문문항 목록
	dim oSurveyQuestion
	Set oSurveyQuestion = new CSurvey

	oSurveyQuestion.FRectSn = srv_sn
	oSurveyQuestion.FPagesize = 15
	oSurveyQuestion.FCurrPage = page
	oSurveyQuestion.FRectUsing = using
	oSurveyQuestion.FRectOrder = "desc"

	oSurveyQuestion.GetSurveyQstList

%>
<script type='text/javascript'>
<!--
	// 페이지 이동
	function goPage(pg)
	{
		document.frm_list.page.value=pg;
		document.frm_list.submit();
	}

	// 문항 등록
	function popQstWrite(ssn) {
		var popSurvey = window.open("survey_qst_write.asp?ssn="+ssn,"QuestPop","width=1200,height=768,scrollbars=yes");
		popSurvey.focus();
	}

	// 문항 수정
	function popQstModify(ssn,qsn) {
		var popSurvey = window.open("survey_qst_modi.asp?ssn="+ssn+"&qsn="+qsn,"QuestPop","width=1200,height=768,scrollbars=yes");
		popSurvey.focus();
	}
//-->
</script>
<!-- 설문정보 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="10%" bgcolor="<%= adminColor("gray") %>">설문번호</td>
	<td width="40%" align="left"><%=srv_sn%></td>
	<td width="10%" bgcolor="<%= adminColor("gray") %>">상태</td>
	<td width="40%" align="left"><%=oSurveyMaster.FitemList(1).getSurveyState%></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">기간</td>
	<td align="left"><%=left(oSurveyMaster.FitemList(1).Fsrv_startDt,10) & " ~ " & left(oSurveyMaster.FitemList(1).Fsrv_endDt,10)%></td>
	<td bgcolor="<%= adminColor("gray") %>">구분</td>
	<td align="left">
	<%
		Select Case oSurveyMaster.FitemList(1).Fsrv_div
			Case "1"
				Response.Write "업체"
			Case "2"
				Response.Write "직원"
		end Select
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">제목</td>
	<td align="left" colspan="3"><%= ReplaceBracket(oSurveyMaster.FitemList(1).Fsrv_subject) %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">머리말</td>
	<td align="left" colspan="3"><%= nl2br(ReplaceBracket(replace(oSurveyMaster.FitemList(1).Fsrv_head,"<","&lt;"))) %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">꼬리말</td>
	<td align="left" colspan="3"><%= nl2br(ReplaceBracket(replace(oSurveyMaster.FitemList(1).Fsrv_tail,"<","&lt;"))) %></td>
</tr>
</table>
<!-- 설문정보 끝 -->
<!-- 설문액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>&nbsp;</td>
	<td align="right" style="padding:4 0 4 0"><input type="button" class="button" value="정보수정" onClick="window.open('survey_write.asp?sn=<%=srv_sn%>','SurveyPop','width=1400,height=768')"></td>
</tr>
</table>
<!-- 설문액션 끝 -->
<!-- 메인 목록 시작 -->
<form name="frm_list" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sn" value="<%=srv_sn%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="div" value="<%=div%>">
<input type="hidden" name="using" value="<%=using%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="6">
		검색결과 : <b><%=FormatNumber(oSurveyQuestion.FTotalCount,0)%></b>
		&nbsp;
		페이지 : <b><%= page %>/<%=FormatNumber(oSurveyQuestion.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>번호</td>
	<td>형태</td>
	<td>문항</td>
	<td>필수여부</td>
	<td>지문</td>
	<td>상태</td>
</tr>
<%
	if oSurveyQuestion.FResultCount=0 then
%>
<tr>
	<td colspan="6" height="60" align="center" bgcolor="#FFFFFF">등록(검색)된 문항이 없습니다.</td>
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

			'사용여부
			if oSurveyQuestion.FitemList(lp).Fqst_isusing="Y" then
				strDel = "<font color=darkblue>사용</font>"
			else
				strDel = "<font color=darkred>삭제</font>"
			end if

%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oSurveyQuestion.FitemList(lp).Fqst_sn%></td>
	<td><%=strType%></td>
	<td><a href="javascript:popQstModify(<%=srv_sn%>,<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>)"><%=oSurveyQuestion.FitemList(lp).Fqst_content%></a></td>
	<td><% if oSurveyQuestion.FitemList(lp).Fqst_isNull="N" then Response.Write "답변필수": Else Response.Write "공란허용": End if %></td>
	<td><%=oSurveyQuestion.FitemList(lp).FpollCnt%></td>
	<td><%=strDel%></td>
</tr>
<%
		next
	end if
%>
<!-- 메인 목록 끝 -->
<!-- 페이지 시작 -->
<tr>
	<td colspan="6" align="center" bgcolor="<%= adminColor("tabletop") %>">
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
</table>
</form>
<!-- 페이지 끝 -->
<!-- 문항액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding:4 0 4 0"><input type="button" class="button" value="미리보기" onClick="window.open('survey_preview.asp?sn=<%=srv_sn%>','PreviewPop','width=778,height=700,scrollbars=yes')"></td>
	<td align="right" style="padding:4 0 4 0"><input type="button" class="button" value="문항등록" onClick="popQstWrite(<%=srv_sn%>)"></td>
</tr>
</table>
<!-- 문항액션 끝 -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->