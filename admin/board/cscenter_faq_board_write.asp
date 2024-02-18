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
고객센타 - 자주묻는질문<br><br>
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
<form method="post" name="f" action="cscenter_faq_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="subcd" value="00">
메뉴 :
<select name="divcd">
  <option value="01">회원정보관련 FAQ</option>
  <option value="02">상품문의 FAQ</option>
  <option value="03">주문/결재 FAQ</option>
  <option value="04">취소/반품 FAQ</option>
  <option value="05">기타 FAQ</option>
</select><br>
제목 : <input type="text" name="title" size="30" value=""><br>
내용 : <textarea name="contents" cols="80" rows="12"></textarea><br><br>

<input type="button" value=" 등록 " onclick="SubmitForm()">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->