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

	'�⺻�� ����
	if page="" then page=1

	'// �������� ����
	dim oSurveyMaster
	Set oSurveyMaster = new CSurvey

	oSurveyMaster.FRectSn = srv_sn
	
	oSurveyMaster.GetSurveyStatistCont

	'// �������� ���
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
	// ������ �̵�
	function goPage(pg)
	{
		document.frm_list.page.value=pg;
		document.frm_list.submit();
	}

	// ��Ÿ�ǰ�,�ְ��� �亯 �˾�
	function popCommentView(qstSn)
	{
		window.open("survey_statist_comment.asp?qsn="+qstSn,"popComView","width=720,height=600,scrollbars=yes");
	}
//-->
</script>
<script language="javascript" src="/lib/util/chart/FusionCharts.js"></script>
<!-- �������� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="10%" bgcolor="<%= adminColor("gray") %>">������ȣ</td>
	<td width="40%" align="left"><%=srv_sn%></td>
	<td width="10%" bgcolor="<%= adminColor("gray") %>">����</td>
	<td width="40%" align="left">
	<%
		if oSurveyMaster.FitemList(1).Fsrv_isusing="Y" then
			if date()<oSurveyMaster.FitemList(1).Fsrv_startDt then
				Response.Write "<font color=darkgreen>���</font>"
			elseif date()>oSurveyMaster.FitemList(1).Fsrv_endDt then
				Response.Write "<font color=darkorange>����</font>"
			else
				Response.Write "<font color=darkblue>������</font>"
			end if
		else
			Response.Write "<font color=darkred>����</font>"
		end if
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">�Ⱓ</td>
	<td align="left"><%=left(oSurveyMaster.FitemList(1).Fsrv_startDt,10) & " ~ " & left(oSurveyMaster.FitemList(1).Fsrv_endDt,10)%></td>
	<td bgcolor="<%= adminColor("gray") %>">������</td>
	<td align="left"><%=oSurveyMaster.FitemList(1).FansCnt%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">����</td>
	<td align="left" colspan="3"><%=oSurveyMaster.FitemList(1).Fsrv_subject%></td>
</tr>
</table>
<!-- �������� �� -->
<!-- �����׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>&nbsp;</td>
</tr>
</table>
<!-- �����׼� �� -->
<!-- ���� ��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_list" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sn" value="<%=srv_sn%>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="div" value="<%=div%>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="5">
		�˻���� : <b><%=FormatNumber(oSurveyQuestion.FTotalCount,0)%></b>
		&nbsp;
		������ : <b><%= page %>/<%=FormatNumber(oSurveyQuestion.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">��ȣ</td>
	<td width="50">����</td>
	<td>����</td>
	<td width="410">�亯���</td>
	<td width="60">��Ÿ</td>
</tr>
<%
	if oSurveyQuestion.FResultCount=0 then
%>
<tr>
	<td colspan="5" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
</tr>
<%
	else
		for lp=0 to oSurveyQuestion.FResultCount - 1
			'����
			Select Case oSurveyQuestion.FitemList(lp).Fqst_type
				Case "1"
					strType = "������"
				Case "2"
					strType = "�ְ���"
				Case "3"
					strType = "�ܴ���"
				Case "9"
					strType = "������"
			end Select
%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oSurveyQuestion.FitemList(lp).Fqst_sn%></td>
	<td><%=strType%></td>
	<td align="left"><%=oSurveyQuestion.FitemList(lp).Fqst_content%></td>
	<% if strType="������" then %>
	<td>
		<div id="chartdiv<%=lp%>" align="center"></div>
		<script type="text/javascript">	
			var chart = new FusionCharts("/lib/util/chart/MSBar2D.swf", "chartdiv<%=lp%>", "400", "150", "0", "0");
			chart.setDataURL("survey_answer_xml.asp?qsn=<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>");
			chart.render("chartdiv<%=lp%>");
		</script>
	</td>
	<td><a href="javascript:popCommentView(<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>)">[�ǰߺ���]</a></td>
	<% elseif strType="�ְ���" then %>
	<td colspan="2"><a href="javascript:popCommentView(<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>)">[�ְ��� �亯 ����]</a></td>
	<% elseif strType="�ܴ���" then %>
	<td colspan="2"><a href="javascript:popCommentView(<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>)">[�ܴ��� �亯 ����]</a></td>
	<% end if %>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<tr>
	<td colspan="5" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!-- ������ ���� -->
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
	<!-- ������ �� -->
	</td>
</tr>
</form>
</table>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->