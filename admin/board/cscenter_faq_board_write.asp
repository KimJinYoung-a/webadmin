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
����Ÿ - ���ֹ�������<br><br>
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
<form method="post" name="f" action="cscenter_faq_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="subcd" value="00">
�޴� :
<select name="divcd">
  <option value="01">ȸ���������� FAQ</option>
  <option value="02">��ǰ���� FAQ</option>
  <option value="03">�ֹ�/���� FAQ</option>
  <option value="04">���/��ǰ FAQ</option>
  <option value="05">��Ÿ FAQ</option>
</select><br>
���� : <input type="text" name="title" size="30" value=""><br>
���� : <textarea name="contents" cols="80" rows="12"></textarea><br><br>

<input type="button" value=" ��� " onclick="SubmitForm()">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->