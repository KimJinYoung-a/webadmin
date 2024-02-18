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
'공지사항
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
고객센타 - 공지사항<br><br>
<script>
function SubmitForm()
{
        if (document.f.title.value == "") {
                alert("제목을 입력하세요.");
                return;
        }
        if (document.f.contents.value == "") {
                alert("내용을 입력하세요.");
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
	<td class="a" align="center" width="120">제목</td>
	<td><input type="text" name="title" size="60" value="<%= boardnotice.results(0).title %>" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td class="a" align="center">내용</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"><%= db2html(boardnotice.results(0).contents) %></textarea><br><font color="red">(실제적용사이즈 입니다. 적당히 엔터키로 줄맞춤해주세요!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td class="a" align="center">사용유무</td>
	<td><input type="radio" name="isusing" value="Y" <% if boardnotice.results(0).isusing = "Y" then response.write "checked" %>>사용 <input type="radio" name="isusing" value="N" <% if boardnotice.results(0).isusing = "N" then response.write "checked" %>>사용안함</td>
</tr>
</form>
</table>
<br>
<input type="button" value=" 수정 " onclick="SubmitForm()">
<br><br>
(링크 양식)<br>

&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;소냐 이벤트 바로가기&lt;/a&gt;

<!-- #include virtual="/lib/db/db2close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->