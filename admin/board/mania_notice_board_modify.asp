<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/db2_manianewscls.asp" -->
<%

dim i, j

'==============================================================================
'��������
dim boardnotice
set boardnotice = New CBoardNotice

boardnotice.read(request("id"))

%>
<STYLE TYPE="text/css">
<!--
    A:link, A:visited, A:active { text-decoration: none; }
    A:hover { text-decoration:underline; }
    BODY, TD, UL, OL, PRE { font-size: 10pt; }
-->
</STYLE>
<link rel=stylesheet type="text/css" href="/bct.css">
����Ÿ - ��������<br><br>
<script>
function SubmitForm()
{
        if (document.f.title.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("������ �Է��ϼ���.");
                return;
        }
        document.f.submit();
}
</script>


<table border="0" cellpadding="0" cellspacing="1" bgcolor="#B0B0B0" class="a">
<form method="post" name="f" action="mania_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= request("id") %>">
<input type="hidden" name="mode" value="modify">
<tr bgcolor="#FFFFFF">
	<td class="a" align="center" width="120">����</td>
	<td><input type="text" name="title" size="60" value="<%= boardnotice.results(0).title %>" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td class="a" align="center">����</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"><%= db2html(boardnotice.results(0).contents) %></textarea><br><font color="red">(������������� �Դϴ�. ������ ����Ű�� �ٸ������ּ���!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td class="a" align="center">�������</td>
	<td><input type="radio" name="isusing" value="Y" <% if boardnotice.results(0).isusing = "Y" then response.write "checked" %>>��� <input type="radio" name="isusing" value="N" <% if boardnotice.results(0).isusing = "N" then response.write "checked" %>>������</td>
</tr>
</form>
</table>
<br>
<input type="button" value=" ���� " onclick="SubmitForm()">
<br><br>
(��ũ ���)<br>

&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;�ҳ� �̺�Ʈ �ٷΰ���&lt;/a&gt;

<!-- #include virtual="/lib/db/db2close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->