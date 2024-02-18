<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/offshop_noticecls.asp" -->
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
    INPUT,SELECT,TEXTAREA { border:1 solid #666666; background-color: #CACACA; color: #000000; }
-->
</STYLE>
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
        if (document.f.yuhyostart.value == "") {
                alert("유효시작일을 입력하세요.");
                return;
        }
        if (document.f.yuhyoend.value == "") {
                alert("유효종료일을 입력하세요.");
                return;
        }

        document.f.submit();
}
function SubmitDelete()
{
        if (confirm("삭제하시겠습니까?") == true) {
                document.f.mode.value = "delete";
                document.f.submit();
        }
}
</script>
<form method="post" name="f" action="offshop_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= request("id") %>">
<input type="hidden" name="mode" value="modify">
몰타입 : <select name="malltype">
				<option value="" <% if boardnotice.results(0).malltype = "" then response.write "selected" %>>선택</option>
				<option value="00" <% if boardnotice.results(0).malltype = "00" then response.write "selected" %>>전체</option>
				<option value="01" <% if boardnotice.results(0).malltype = "01" then response.write "selected" %>>1F Shop</option>
				<option value="02" <% if boardnotice.results(0).malltype = "02" then response.write "selected" %>>1F Cafe</option>
				<option value="03" <% if boardnotice.results(0).malltype = "03" then response.write "selected" %>>3F Zoom</option>
				<option value="04" <% if boardnotice.results(0).malltype = "04" then response.write "selected" %>>3F College</option>
			</select><br>
공지유형 : <select name="noticetype">
				<option value="" <% if boardnotice.results(0).noticetype = "" then response.write "selected" %>>선택</option>
				<option value="01" <% if boardnotice.results(0).noticetype = "01" then response.write "selected" %>>전체공지</option>
				<option value="02" <% if boardnotice.results(0).noticetype = "02" then response.write "selected" %>>제품공지</option>
				<option value="03" <% if boardnotice.results(0).noticetype = "03" then response.write "selected" %>>이벤트공지</option>
			</select><br>
제목 : <input type="text" size="30" name="title" value="<%= boardnotice.results(0).title %>"><br>
내용 : <textarea name="contents" cols="80" rows="6"><%= db2html(boardnotice.results(0).contents) %></textarea><br>
유효시작일 : <input type="text" name="yuhyostart" value="<%= boardnotice.results(0).yuhyostart %>"><br>
유효종료일 : <input type="text" name="yuhyoend" value="<%= boardnotice.results(0).yuhyoend %>"><br><br>

<input type="button" value=" 수정 " onclick="SubmitForm()">
<input type="button" value=" 삭제 " onclick="SubmitDelete()">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->