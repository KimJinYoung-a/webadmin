<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<%
	Dim vECode, vQuery, vFB_appid, vFB_content
	vECode = Request("ecode")
	If vECode = "" Then
		Response.End
	End If
	
	vQuery = "select fb_appid, fb_content from [db_event].[dbo].[tbl_event_display] where evt_code = '" & vECode & "'"
	rsget.Open vQuery,dbget,1
	If Not rsget.Eof Then
		vFB_appid	= rsget("fb_appid")
		vFB_content	= db2html(rsget("fb_content"))
	End IF
	rsget.Close
%>

<Script language="javascript">
function gofbcontentsave()
{
	if(frm.fb_appid.value == "")
	{
		alert("�� ID�� �Է��ϼ���.");
		frm.fb_appid.focus();
		return;
	}
	if(frm.fb_content.value == "")
	{
		alert("����html�� �Է��ϼ���.");
		frm.fb_content.focus();
		return;
	}
	frm.submit();
}
</script>

<table cellpadding="0" cellspacing="0" class="a">
<tr>
	<td><b><font size="2">���̽��� �� ���� �Է�â</font></b></td>
</tr>
</table>

<table width="100%" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" methopd="post" action="pop_event_facebookapp_proc.asp">
<input type="hidden" name="ecode" value="<%=vECode%>">
<tr>
	<td align="center" width="70" bgcolor="<%= adminColor("tabletop") %>">�� ID</td>
	<td bgcolor="#FFFFFF"><input type="text" name="fb_appid" value="<%=vFB_appid%>" size="50"></td>
</tr>
<tr>
	<td align="center" width="70" bgcolor="<%= adminColor("tabletop") %>">����html</td>
	<td bgcolor="#FFFFFF">
		�� �������� �̹��� ������� <b>����</b> �ִ� ����� <b>520px</b> �Դϴ�.<br>
		<textarea name="fb_content" cols="53" rows="18"><%=vFB_content%></textarea>
	</td>
</tr>
<tr>
	<td bgcolor="#FFFFFF" colspan="2" align="right"><input type="button" class="button" value="��  ��" onClick="gofbcontentsave()"></td>
</tr>
</form>
</table>

<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->