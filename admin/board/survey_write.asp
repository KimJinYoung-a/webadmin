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
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/board/surveyCls.asp" -->
<%
	Dim srv_sn, srv_subject, srv_startDt, srv_endDt, srv_head, srv_tail, srv_div, mode

	srv_sn = Request("sn")

	if srv_sn<>"" then
		mode = "srv_edit"

		'// �������� ����
		dim oSurveyMaster
		Set oSurveyMaster = new CSurvey
		oSurveyMaster.FRectSn = srv_sn
		oSurveyMaster.GetSurveyCont

		srv_subject	= ReplaceBracket(oSurveyMaster.FitemList(1).Fsrv_subject)
		srv_div		= oSurveyMaster.FitemList(1).Fsrv_div
		srv_startDt	= oSurveyMaster.FitemList(1).Fsrv_startDt
		srv_endDt	= oSurveyMaster.FitemList(1).Fsrv_endDt
		srv_head	= ReplaceBracket(oSurveyMaster.FitemList(1).Fsrv_head)
		srv_tail	= ReplaceBracket(oSurveyMaster.FitemList(1).Fsrv_tail)
	else
		mode = "srv_add"
		srv_startDt	= date()
		srv_endDt	= dateadd("d",15,date())
	end if
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type='text/javascript'>
<!--
	function chkSubmit() {
		var frm = document.frm;

		if(!frm.srv_subject.value) {
			alert("������ ������ �Է����ּ���.\n\n�������� ��������â�� �������� ���Դϴ�.");
			frm.srv_subject.focus();
			return false;
		}

		if(!frm.srv_div.value) {
			alert("������ ������ �������ּ���.");
			frm.srv_div.focus();
			return false;
		}

		if(!frm.srv_startDt.value||!frm.srv_endDt.value) {
			alert("���� ������ �� �������� �������ּ���.");
			frm.srv_startDt.focus();
			return false;
		}

		if(frm.srv_startDt.value>=frm.srv_endDt.value) {
			alert("�������� �����Ϻ��� ���ų� �ʽ��ϴ�.\n�Ⱓ�� Ȯ�����ּ���.");
			frm.srv_startDt.focus();
			return false;
		}

		// ������
		return true;
	}

	function closeWin() {
		if(confirm("���� ������ ����Ͻðڽ��ϱ�?")) {
			self.close();
		}
	}
//-->
</script>
<form name="frm" method="POST" action="survey_process.asp" onSubmit="return chkSubmit()">
<input type="hidden" name="mode" value="<%=mode%>">
<input type="hidden" name="srv_sn" value="<%=srv_sn%>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td colspan="2" bgcolor="<%= adminColor("sky") %>" align="left"><img src="/images/icon_star.gif" align="absmiddle"><b> <%if srv_sn="" then Response.Write "���� �űԵ��": Else Response.Write "���� ��������": end if%></b></td>
</tr>
<% if srv_sn<>"" then %>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="20%" bgcolor="<%= adminColor("gray") %>">���� ��ȣ</td>
	<td width="80%" align="left"><b><%=srv_sn%></b></td>
</tr>
<% end if %>
<tr align="center" bgcolor="#FFFFFF" >
	<td width="20%" bgcolor="<%= adminColor("gray") %>">���� ����</td>
	<td width="80%" align="left"><input type="text" name="srv_subject" class="text" size="80" value="<%=srv_subject%>"></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">��� ����</td>
	<td align="left">
		<select name="srv_div" class="select">
			<option value="">::���м���::</option>
			<option value="1">��ü����</option>
			<option value="2">��������</option>
		</select>
		<script language="javascript">
		frm.srv_div.value="<%=srv_div%>";
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">���� �Ⱓ</td>
	<td align="left">
        <input id="srv_startDt" name="srv_startDt" value="<%=srv_startDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="srv_startDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
        <input id="srv_endDt" name="srv_endDt" value="<%=srv_endDt%>" class="text" size="10" maxlength="10" />
        <img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="srv_endDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "srv_startDt", trigger    : "srv_startDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "srv_endDt", trigger    : "srv_endDt_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">���� �Ӹ���</td>
	<td align="left"><textarea name="srv_head" class="textarea" style="width:100%; height:100px;"><%=srv_head%></textarea></td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td bgcolor="<%= adminColor("gray") %>">���� ������</td>
	<td align="left"><textarea name="srv_tail" class="textarea" style="width:100%; height:100px;"><%=srv_tail%></textarea></td>
</tr>
<tr bgcolor="#FFFFFF" >
	<td colspan="2" align="center">
		<input type="submit" value=" �� �� " class="button"> &nbsp; &nbsp;
		<input type="button" value=" �� �� " class="button" onClick="closeWin();">
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->