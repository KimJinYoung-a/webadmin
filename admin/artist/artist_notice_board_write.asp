<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2012.03.22 ������ �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 10pt; }
-->
</STYLE>
<link rel=stylesheet type="text/css" href="/bct.css">
<script>
function SubmitForm()
{
    if (document.f.title.value == "") {
        alert("������ �Է��ϼ���.");
        document.f.title.focus();
        return;
    }
    if (document.f.contents.value == "") {
        alert("������ �Է��ϼ���.");
        document.f.contents.focus();
        return;
    }
	if (confirm("����Ͻðڽ��ϱ�?") == true) {
		document.f.submit();
	}
}
</script>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080" class="a">
<form method="post" name="f" action="artist_notice_board_process.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
<tr bgcolor="#FFFFFF">
	<td>����</td>
	<td><input type="text" name="title" size="60" value="" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"></textarea><br><font color="red">(������������� �Դϴ�. ������ ����Ű�� �ٸ������ּ���!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����������</td>
	<td><input type="radio" name="fixyn" value="Y">��� <input type="radio" name="fixyn" value="N" checked>������</td>
</tr>
</form>
</table>
<br>
<input type="button" value=" ��� " onclick="SubmitForm()">
<br><br>
(��ũ ���)<br>
&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;�ҳ� �̺�Ʈ �ٷΰ���&lt;/a&gt;

<!-- #include virtual="/lib/db/dbclose.asp" -->