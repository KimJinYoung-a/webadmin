<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/board/lib/classes/offshop_noticecls.asp" -->
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
function SubmitDelete()
{
        if (confirm("�����Ͻðڽ��ϱ�?") == true) {
                document.f.mode.value = "delete";
                document.f.submit();
        }
}
</script>
<form method="post" name="f" action="offshop_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="id" value="<%= request("id") %>">
<input type="hidden" name="mode" value="modify">
��Ÿ�� : <select name="malltype">
				<option value="" <% if boardnotice.results(0).malltype = "" then response.write "selected" %>>����</option>
				<option value="00" <% if boardnotice.results(0).malltype = "00" then response.write "selected" %>>��ü</option>
				<option value="01" <% if boardnotice.results(0).malltype = "01" then response.write "selected" %>>1F Shop</option>
				<option value="02" <% if boardnotice.results(0).malltype = "02" then response.write "selected" %>>1F Cafe</option>
				<option value="03" <% if boardnotice.results(0).malltype = "03" then response.write "selected" %>>3F Zoom</option>
				<option value="04" <% if boardnotice.results(0).malltype = "04" then response.write "selected" %>>3F College</option>
			</select><br>
�������� : <select name="noticetype">
				<option value="" <% if boardnotice.results(0).noticetype = "" then response.write "selected" %>>����</option>
				<option value="01" <% if boardnotice.results(0).noticetype = "01" then response.write "selected" %>>��ü����</option>
				<option value="02" <% if boardnotice.results(0).noticetype = "02" then response.write "selected" %>>��ǰ����</option>
				<option value="03" <% if boardnotice.results(0).noticetype = "03" then response.write "selected" %>>�̺�Ʈ����</option>
			</select><br>
���� : <input type="text" size="30" name="title" value="<%= boardnotice.results(0).title %>"><br>
���� : <textarea name="contents" cols="80" rows="6"><%= db2html(boardnotice.results(0).contents) %></textarea><br>
��ȿ������ : <input type="text" name="yuhyostart" value="<%= boardnotice.results(0).yuhyostart %>"><br>
��ȿ������ : <input type="text" name="yuhyoend" value="<%= boardnotice.results(0).yuhyoend %>"><br><br>

<input type="button" value=" ���� " onclick="SubmitForm()">
<input type="button" value=" ���� " onclick="SubmitDelete()">
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->