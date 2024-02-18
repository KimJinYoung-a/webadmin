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

	'�⺻�� ����
	if page="" then page=1

	'// ���� ����
	dim oSurveyQuest
	Set oSurveyQuest = new CSurvey
	oSurveyQuest.FRectSn = qst_sn
	oSurveyQuest.GetSurveyQuestCont

	'// �ְ��� ���
	dim oSurvey
	Set oSurvey = new CSurvey

	oSurvey.FRectSn = qst_sn
	oSurvey.FPagesize = 15
	oSurvey.FCurrPage = page

	oSurvey.GetSurveyCommentList
%>
<script language="javascript">
<!--
	// ������ �̵�
	function goPage(pg)
	{
		document.frm_list.page.value=pg;
		document.frm_list.submit();
	}
//-->
</script>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr><td><b>�� �ְ��� �ǰ� ����</b></td></tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="10%" bgcolor="<%= adminColor("gray") %>">������ȣ</td>
	<td width="40%" align="left"><%=qst_sn%></td>
	<td width="10%" bgcolor="<%= adminColor("gray") %>">�ʼ�����</td>
	<td width="40%" align="left">
	<%
		if oSurveyQuest.FitemList(1).Fqst_isNull="Y" then
			Response.Write "<font color=darkblue>�������</font>"
		else
			Response.Write "<font color=darkred>�亯�ʼ�</font>"
		end if
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">����</td>
	<td align="left" colspan="3"><%=oSurveyQuest.FitemList(1).Fqst_content%></td>
</tr>
</table>
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr><td>&nbsp;</td></tr>
</table>
<!-- ���� ��� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm_list" method="get" action="">
<input type="hidden" name="qsn" value="<%=qst_sn%>">
<input type="hidden" name="page" value="">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="3">
		��� : <b><%=FormatNumber(oSurvey.FTotalCount,0)%></b>
		&nbsp;
		������ : <b><%= page %>/<%=FormatNumber(oSurvey.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="40">��ȣ</td>
	<% if oSurveyQuest.FitemList(1).Fqst_type="1" then %>
		<td width="160">��������</td>
		<td>����</td>
	<% else %>
		<td width="660">����</td>
	<% end if %>
	
</tr>
<%
	if oSurvey.FResultCount=0 then
%>
<tr>
	<td colspan="3" height="60" align="center" bgcolor="#FFFFFF">��ϵ� ������ �����ϴ�.</td>
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
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<tr>
	<td colspan="3" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<!-- ������ ���� -->
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
	<!-- ������ �� -->
	</td>
</tr>
</form>
</table>
<!-- ������ �� -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->