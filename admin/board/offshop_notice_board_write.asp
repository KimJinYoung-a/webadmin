<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
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
</script>
<form method="post" name="f" action="offshop_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
몰타입 : <select name="malltype">
				<option value="">선택</option>
				<option value="00">전체</option>
				<option value="01">1F Shop</option>
				<option value="02">1F Cafe</option>
				<option value="03">2F Zoom</option>
				<option value="04">3F College</option>
			</select><br>
공지유형 : <select name="noticetype">
				<option value="">선택</option>
				<option value="01">전체공지</option>
				<option value="02">제품공지</option>
				<option value="03">이벤트공지</option>
			</select><br>
제목 : <input type="text" name="title" size="30" value=""><br>
내용 : <textarea name="contents" cols="80" rows="6"></textarea><br>
유효시작일 : <input type="text" name="yuhyostart" value=""><br>
유효종료일 : <input type="text" name="yuhyoend" value=""><br><br>

<input type="button" value=" 등록 " onclick="SubmitForm()">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->