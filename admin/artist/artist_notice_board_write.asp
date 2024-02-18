<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2012.03.22 김진영 추가
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
        alert("제목을 입력하세요.");
        document.f.title.focus();
        return;
    }
    if (document.f.contents.value == "") {
        alert("내용을 입력하세요.");
        document.f.contents.focus();
        return;
    }
	if (confirm("등록하시겠습니까?") == true) {
		document.f.submit();
	}
}
</script>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080" class="a">
<form method="post" name="f" action="artist_notice_board_process.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
<tr bgcolor="#FFFFFF">
	<td>제목</td>
	<td><input type="text" name="title" size="60" value="" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>내용</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"></textarea><br><font color="red">(실제적용사이즈 입니다. 적당히 엔터키로 줄맞춤해주세요!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>고정글유무</td>
	<td><input type="radio" name="fixyn" value="Y">사용 <input type="radio" name="fixyn" value="N" checked>사용안함</td>
</tr>
</form>
</table>
<br>
<input type="button" value=" 등록 " onclick="SubmitForm()">
<br><br>
(링크 양식)<br>
&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;소냐 이벤트 바로가기&lt;/a&gt;

<!-- #include virtual="/lib/db/dbclose.asp" -->