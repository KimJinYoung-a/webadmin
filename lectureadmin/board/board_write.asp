<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/board_cls.asp"-->
<%
	'// ���� ���� //
	dim oBoard

	'// ���� ����
	set oBoard = new Cboard
%>
<script language='javascript'>
<!--
	// �Է��� �˻�
	function chk_form(frm)
	{
		if(!frm.qstTitle.value)
		{
			alert("������ �Է����ֽʽÿ�.");
			frm.qstTitle.focus();
			return false;
		}

		if(!frm.qstCont.value)
		{
			alert("������ �ۼ����ֽʽÿ�.");
			frm.qstCont.focus();
			return false;
		}

		// �� ����
		return true;
	}
//-->
</script>
<!-- ���� ȭ�� ���� -->
<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
<form name="frm_write" method="POST" onSubmit="return chk_form(this)" action="doBoard.asp">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr align="center" bgcolor="#F0F0FD">
	<td height="26" align="left" colspan="2"><b>�Խù� �ű� ���</b></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF">
		<select name="commCd">
		<%=oBoard.optCommCd("'G000'", "")%>
		</select>
	</td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><input type="text" name="qstTitle" size="40" maxlength="40"></td>
</tr>
<tr>
	<td align="center" width="120" bgcolor="#DDDDFF"><font color="darkred">* </font>����</td>
	<td bgcolor="#FFFFFF"><textarea name="qstCont" rows="14" cols="80"></textarea></td>
</tr>
<tr><td height="1" colspan="2" bgcolor="#D0D0D0"></td></tr>
<tr>
	<td colspan="2" height="32" bgcolor="#FAFAFA" align="center">
		<input type="image" src="/images/icon_save.gif" style="border:0px;cursor:pointer" align="absmiddle"> &nbsp;
		<img src="/images/icon_cancel.gif" onClick="history.back()" style="cursor:pointer" align="absmiddle">
	</td>
</tr>
</form>
</table>
<%	set oBoard = Nothing %>
<!-- ���� ȭ�� �� -->
<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbacademyclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
