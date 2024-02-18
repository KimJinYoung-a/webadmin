<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��������
' Hieditor : ������ ����
'			 2022.07.08 �ѿ�� ����(isms�����������ġ, ǥ���ڵ�κ���)
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

	'�⺻�� ����
	if page="" then page=1
	if using="" then using="Y"


	'// ���� ���
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
	// ������ �̵�
	function goPage(pg)
	{
		document.frm.page.value=pg;
		document.frm.submit();
	}
//-->
</script>
<!-- ��� �˻��� ���� -->
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="80" bgcolor="<%= adminColor("gray") %>">�˻�����</td>
	<td align="left">
		���� <select name="div" class="select">
			<option value="">��ü</option>
			<option value="1">��ü</option>
			<option value="2">����</option>
		</select>
		/ ���� <select name="using" class="select">
			<option value="N">����</option>
			<option value="Y">���</option>
		</select>
		<script language="javascript">
			document.frm.div.value="<%=div%>";
			document.frm.using.value="<%=using%>";
		</script>
	</td>
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="submit" class="button_s" value="�˻�">
	</td>
</tr>
</table>
</form>
<br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>&nbsp;</td>
	<td align="right" style="padding:4 0 4 0"><input type="button" class="button" value="�űԵ��" onClick="window.open('survey_write.asp','SurveyPop','width=1400,height=768')"></td>
</tr>
</table>
<!-- �׼� �� -->
<!-- ���� ��� ���� -->
<form name="frm_list" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="div" value="<%=div%>">
<input type="hidden" name="using" value="<%=using%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="8">
		�˻���� : <b><%=FormatNumber(oSurvey.FTotalCount,0)%></b>
		&nbsp;
		������ : <b><%= page %>/<%=FormatNumber(oSurvey.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>�Ϸù�ȣ</td>
	<td>��������</td>
	<td>����</td>
	<td>������</td>
	<td>������</td>
	<td>�����</td>
	<td>����</td>
</tr>
<%
	if oSurvey.FResultCount=0 then
%>
<tr>
	<td colspan="7" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
</tr>
<%
	else
		for lp=0 to oSurvey.FResultCount - 1
			'����
			Select Case oSurvey.FitemList(lp).Fsrv_div
				Case "1"
					strDiv = "��ü"
				Case "2"
					strDiv = "����"
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
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<tr>
	<td colspan="7" align="center" bgcolor="<%= adminColor("tabletop") %>">
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
</table>
</form>
<!-- ������ �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->