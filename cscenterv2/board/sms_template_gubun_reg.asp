<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/board/cs_templatecls.asp"-->
<%
dim mode
dim idx, mastergubun, gubun, gubunname, contents, disporder, isusing


mode = RequestCheckvar(request("mode"),16)

mastergubun = "10"		'// SMS
idx = RequestCheckvar(request("idx"),10)


dim osmstemplate
set osmstemplate = New CCSTemplate
osmstemplate.FCurrPage = 1
osmstemplate.FPageSize = 1
osmstemplate.FRectIdx = idx
osmstemplate.FRectMasterGubun = mastergubun

if (mode <> "addgubun") then
	osmstemplate.GetCSTemplateList

	gubun		= osmstemplate.FItemList(0).Fgubun
	gubunname	= osmstemplate.FItemList(0).Fgubunname
	contents	= osmstemplate.FItemList(0).Fcontents
	disporder	= osmstemplate.FItemList(0).Fdisporder
	isusing		= osmstemplate.FItemList(0).Fisusing
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

	if (confirm('���� �Ͻðڽ��ϱ�?') == true) {
		frm.submit();
	}
}

//-->
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form method="post" name="frm" action="sms_template_process.asp">
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
		[���̵�] : ���� ���̵�<br>
		[�̸�] : �ۼ��� �̸�<br><br>

		*��ü����<br>
		[��ü��ǰ�ּ�] : ��ü ��ǰ�ּ�<br>
		[��ü��ǰ�����] : ��ü ��ǰ�����<br>
		[��ü��ǰ��ȭ] : ��ü��ǰ��ȭ<br>
		[��ü�ŷ��ù��] : ��ü �ŷ��ù��<br>
		[��ü��Ʈ��Ʈ��] : ��ü ��Ʈ��Ʈ��
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

<% set osmstemplate = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->