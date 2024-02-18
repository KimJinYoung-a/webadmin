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
	Dim page, lp, div, using, strType, strDel, srv_sn
	srv_sn = Request("sn")
	page = requestCheckVar(getNumeric(request("page")),10)
	div = Request("div")
	using = requestCheckVar(request("using"),1)

	'�⺻�� ����
	if page="" then page=1
	if using="" then using="Y"

	'// �������� ����
	dim oSurveyMaster
	Set oSurveyMaster = new CSurvey
	oSurveyMaster.FRectSn = srv_sn
	oSurveyMaster.GetSurveyCont

	'// �������� ���
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
	// ������ �̵�
	function goPage(pg)
	{
		document.frm_list.page.value=pg;
		document.frm_list.submit();
	}

	// ���� ���
	function popQstWrite(ssn) {
		var popSurvey = window.open("survey_qst_write.asp?ssn="+ssn,"QuestPop","width=1200,height=768,scrollbars=yes");
		popSurvey.focus();
	}

	// ���� ����
	function popQstModify(ssn,qsn) {
		var popSurvey = window.open("survey_qst_modi.asp?ssn="+ssn+"&qsn="+qsn,"QuestPop","width=1200,height=768,scrollbars=yes");
		popSurvey.focus();
	}
//-->
</script>
<!-- �������� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="10%" bgcolor="<%= adminColor("gray") %>">������ȣ</td>
	<td width="40%" align="left"><%=srv_sn%></td>
	<td width="10%" bgcolor="<%= adminColor("gray") %>">����</td>
	<td width="40%" align="left"><%=oSurveyMaster.FitemList(1).getSurveyState%></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">�Ⱓ</td>
	<td align="left"><%=left(oSurveyMaster.FitemList(1).Fsrv_startDt,10) & " ~ " & left(oSurveyMaster.FitemList(1).Fsrv_endDt,10)%></td>
	<td bgcolor="<%= adminColor("gray") %>">����</td>
	<td align="left">
	<%
		Select Case oSurveyMaster.FitemList(1).Fsrv_div
			Case "1"
				Response.Write "��ü"
			Case "2"
				Response.Write "����"
		end Select
	%>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">����</td>
	<td align="left" colspan="3"><%= ReplaceBracket(oSurveyMaster.FitemList(1).Fsrv_subject) %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">�Ӹ���</td>
	<td align="left" colspan="3"><%= nl2br(ReplaceBracket(replace(oSurveyMaster.FitemList(1).Fsrv_head,"<","&lt;"))) %></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">������</td>
	<td align="left" colspan="3"><%= nl2br(ReplaceBracket(replace(oSurveyMaster.FitemList(1).Fsrv_tail,"<","&lt;"))) %></td>
</tr>
</table>
<!-- �������� �� -->
<!-- �����׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td>&nbsp;</td>
	<td align="right" style="padding:4 0 4 0"><input type="button" class="button" value="��������" onClick="window.open('survey_write.asp?sn=<%=srv_sn%>','SurveyPop','width=1400,height=768')"></td>
</tr>
</table>
<!-- �����׼� �� -->
<!-- ���� ��� ���� -->
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
		�˻���� : <b><%=FormatNumber(oSurveyQuestion.FTotalCount,0)%></b>
		&nbsp;
		������ : <b><%= page %>/<%=FormatNumber(oSurveyQuestion.FtotalPage,0)%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td>��ȣ</td>
	<td>����</td>
	<td>����</td>
	<td>�ʼ�����</td>
	<td>����</td>
	<td>����</td>
</tr>
<%
	if oSurveyQuestion.FResultCount=0 then
%>
<tr>
	<td colspan="6" height="60" align="center" bgcolor="#FFFFFF">���(�˻�)�� ������ �����ϴ�.</td>
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

			'��뿩��
			if oSurveyQuestion.FitemList(lp).Fqst_isusing="Y" then
				strDel = "<font color=darkblue>���</font>"
			else
				strDel = "<font color=darkred>����</font>"
			end if

%>
<tr align="center" bgcolor="#FFFFFF">
	<td><%=oSurveyQuestion.FitemList(lp).Fqst_sn%></td>
	<td><%=strType%></td>
	<td><a href="javascript:popQstModify(<%=srv_sn%>,<%=oSurveyQuestion.FitemList(lp).Fqst_sn%>)"><%=oSurveyQuestion.FitemList(lp).Fqst_content%></a></td>
	<td><% if oSurveyQuestion.FitemList(lp).Fqst_isNull="N" then Response.Write "�亯�ʼ�": Else Response.Write "�������": End if %></td>
	<td><%=oSurveyQuestion.FitemList(lp).FpollCnt%></td>
	<td><%=strDel%></td>
</tr>
<%
		next
	end if
%>
<!-- ���� ��� �� -->
<!-- ������ ���� -->
<tr>
	<td colspan="6" align="center" bgcolor="<%= adminColor("tabletop") %>">
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
</table>
</form>
<!-- ������ �� -->
<!-- ���׾׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td style="padding:4 0 4 0"><input type="button" class="button" value="�̸�����" onClick="window.open('survey_preview.asp?sn=<%=srv_sn%>','PreviewPop','width=778,height=700,scrollbars=yes')"></td>
	<td align="right" style="padding:4 0 4 0"><input type="button" class="button" value="���׵��" onClick="popQstWrite(<%=srv_sn%>)"></td>
</tr>
</table>
<!-- ���׾׼� �� -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->