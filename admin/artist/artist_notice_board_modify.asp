<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2012.03.22 김진영 추가
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/classes/artist/artist_noticeCls.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i, j
Dim boardnotice

dim SearchKey, SearchString, param, page, idx
page = Request("page")
idx = Request("idx")
SearchKey = Request("SearchKey")
SearchString = Request("SearchString")
menupos = Request("menupos")
param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&menupos=" & menupos
param = "&SearchKey=" & SearchKey & "&SearchString=" & Server.URLencode(SearchString) & "&menupos=" & menupos

Set boardnotice = New CBoardNotice
boardnotice.read(idx)
%>
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
	if (confirm("수정하시겠습니까?") == true) {
		document.f.submit();
	}
}
</script>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080" class="a">
<form method="post" name="f" action="artist_notice_board_process.asp" onsubmit="return false">
<input type="hidden" name="idx" value="<%= idx %>">
<input type="hidden" name="mode" value="modify">
<input type="hidden" name="page" value="<%=page%>">
<input type="hidden" name="SearchKey" value="<%=SearchKey%>">
<input type="hidden" name="SearchString" value="<%=SearchString%>">
<input type="hidden" name="menupos" value="<%=menupos%>">
<tr bgcolor="#FFFFFF">
	<td>제목</td>
	<td><input type="text" name="title" size="60" value="<%= boardnotice.results(0).Ftitle %>" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>내용</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"><%= db2html(boardnotice.results(0).Fcontents) %></textarea><br><font color="red">(실제적용사이즈 입니다. 적당히 엔터키로 줄맞춤해주세요!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>고정글유무</td>
	<td><input type="radio" name="fixyn" value="Y" <% if boardnotice.results(0).Ffixyn = "Y" then response.write "checked" %>>사용 <input type="radio" name="fixyn" value="N" <% if boardnotice.results(0).Ffixyn = "N" then response.write "checked" %>>사용안함</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>사용유무</td>
	<td><input type="radio" name="isusing" value="Y" <% if boardnotice.results(0).Fisusing = "Y" then response.write "checked" %>>사용 <input type="radio" name="isusing" value="N" <% if boardnotice.results(0).Fisusing = "N" then response.write "checked" %>>사용안함</td>
</tr>
</form>
</table>
<br>
<a href="javascript:SubmitForm()" onfocus="this.blur()"><img src="/images/icon_modify.gif" border="0" align="absmiddle"></a>
<a href="artist_notice_board_list.asp?page=<%=page & param%>" onfocus="this.blur()"><img src="/images/icon_list.gif" border="0" align="absmiddle"></a>
<br><br>
<table cellpadding="5" cellspacing="0" border="0" class="a">
<tr>
	<td bgcolor="#F8F8FA" style="border:1px solid #D8D8DA">
		<b>(링크 양식)</b><br>
		&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;소냐 이벤트 바로가기&lt;/a&gt;
	</td>
</tr>
</table>
<!-- #include virtual="/lib/db/dbclose.asp" -->