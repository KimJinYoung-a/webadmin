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
        if (document.f.yuhyostart.value == "") {
                alert("��ȿ�������� �Է��ϼ���.");
                return;
        }
        if (document.f.yuhyoend.value == "") {
                alert("��ȿ�������� �Է��ϼ���.");
                return;
        }

        document.f.submit();
}
</script>
<form method="post" name="f" action="offshop_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
��Ÿ�� : <select name="malltype">
				<option value="">����</option>
				<option value="00">��ü</option>
				<option value="01">1F Shop</option>
				<option value="02">1F Cafe</option>
				<option value="03">2F Zoom</option>
				<option value="04">3F College</option>
			</select><br>
�������� : <select name="noticetype">
				<option value="">����</option>
				<option value="01">��ü����</option>
				<option value="02">��ǰ����</option>
				<option value="03">�̺�Ʈ����</option>
			</select><br>
���� : <input type="text" name="title" size="30" value=""><br>
���� : <textarea name="contents" cols="80" rows="6"></textarea><br>
��ȿ������ : <input type="text" name="yuhyostart" value=""><br>
��ȿ������ : <input type="text" name="yuhyoend" value=""><br><br>

<input type="button" value=" ��� " onclick="SubmitForm()">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->