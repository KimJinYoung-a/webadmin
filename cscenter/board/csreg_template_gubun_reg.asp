<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/board/cs_templatecls.asp"-->
<%
dim mode
dim idx, mastergubun, gubun, gubunname, contents, disporder, isusing


mode = request("mode")
mastergubun = request("mastergubun")

if (mastergubun = "") then
	mastergubun = "30"		'// CS����
end if

idx = request("idx")


dim ocsregtemplate
set ocsregtemplate = New CCSTemplate
ocsregtemplate.FCurrPage = 1
ocsregtemplate.FPageSize = 1
ocsregtemplate.FRectIdx = idx
ocsregtemplate.FRectMasterGubun = mastergubun

if (mode <> "addgubun") then
	ocsregtemplate.GetCSTemplateList

	gubun		= ocsregtemplate.FItemList(0).Fgubun
	gubunname	= ocsregtemplate.FItemList(0).Fgubunname
	contents	= ocsregtemplate.FItemList(0).Fcontents
	disporder	= ocsregtemplate.FItemList(0).Fdisporder
	isusing		= ocsregtemplate.FItemList(0).Fisusing
end if

%>
<script language="JavaScript">
<!--

function SubmitAction(frm) {
	/*
	if (frm.gubun.value == "") {
		alert("������ �Է��ϼ���");
		frm.gubun.focus();
		return;
	}

	if ((frm.gubun.value.length != 2) || (frm.gubun.value*0 != 0)) {
		alert("������ 2������ ���ڸ� �����մϴ�.");
		frm.gubun.focus();
		return;
	}
	*/

	if (frm.gubunname.value == "") {
		alert("���и��� �Է��ϼ���");
		frm.gubunname.focus();
		return;
	}

	if (frm.gubunname.value.length > 15) {
		alert("���и��� 15���ڱ��� �����մϴ�.");
		frm.gubunname.focus();
		return;
	}

	if (frm.disporder.value == "") {
		alert("ǥ�ü����� �Է��ϼ���");
		frm.disporder.focus();
		return;
	}

	if (frm.disporder.value*0 != 0) {
		alert("ǥ�ü����� ���ڸ� �����մϴ�.");
		frm.disporder.focus();
		return;
	}

	if (confirm("���� �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="post" name="frm" action="csreg_template_process.asp">
<input type="hidden" name="menupos" value="<% = menupos %>">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="mastergubun" value="<% = mastergubun %>">
<input type="hidden" name="idx" value="<% = idx %>">

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		����
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text_ro" name="gubun" size="4" value="<%= gubun %>">
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		���и�
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="gubunname" size="30" value="<%= gubunname %>">
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		����
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<textarea name="contents" class="textarea" cols="52" rows="10"><%= contents %></textarea>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		ǥ�ü���
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<input type="text" class="text" name="disporder" size="4" value="<%= disporder %>">
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
	<td class="a" width="80">
		���
	</td>
	<td bgcolor="#FFFFFF" align="left">
		<select name="isusing" class="select">
			<option value="Y">�����</option>
			<option value="N" <% if (isusing = "N") then %>selected<% end if %>>������</option>
		</select>
	</td>
</tr>

<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="40">
	<td class="a" width="80">
		�ڵ���ȯ
	</td>
	<td bgcolor="#FFFFFF" align="left">
		* �Ϲ�����<br>
		[�̸�] : �ۼ��� �̸�<br>
		[������ȭ] : �ۼ��� ������ȭ
	</td>
</tr>

<tr align="left" bgcolor="<%= adminColor("tabletop") %>" height="35">
	<td colspan="2" bgcolor="#FFFFFF">
		&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="button" value="�����ϱ�" onclick="SubmitAction(frm);" class="button">
		<input type="button" value="����ϱ�" onclick="history.back();" class="button">
	</td>
</tr>
</form>
</table>

<% set ocsregtemplate = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
