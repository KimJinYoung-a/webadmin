<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' History : 2008.04.29 �ѿ�� �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
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
function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}

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

		if (confirm("����Ͻðڽ��ϱ�?") == true) {
			document.f.submit();
		}
}
</script>
<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080" class="a">
<form method="post" name="f" action="cscenter_notice_board_act.asp" onsubmit="return false">
<input type="hidden" name="mode" value="write">
<tr bgcolor="#FFFFFF">
	<td>��������</td>
	<td>
		  <select name="noticetype">
				<option value="">����</option>
				<!--<option value="01">��ü����</option> 2015�����󿡼� ����. �̻��ش븮.//-->
				<option value="02">�ȳ�</option>
				<option value="03">�̺�Ʈ����</option>
				<option value="04">��۰���</option>
				<option value="05">��÷�ڰ���</option>
				<option value="06">CultureStation</option>
		  </select>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>ī�װ�</td>
	<td>
		<%DrawSelectBoxCategoryOnlyLarge"malltype", "","" %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����</td>
	<td><input type="text" name="title" size="60" value="" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����</td>
	<td><textarea name="contents" cols="70" rows="15" class="textarea2"></textarea><br><font color="red">(������������� �Դϴ�. ������ ����Ű�� �ٸ������ּ���!!)</font></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��ȿ������</td>
	<td><input type="text" size="10" name="yuhyostart" value="" onClick="jsPopCal('f','yuhyostart');" style="cursor:hand;" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>��ȿ������</td>
	<td><input type="text" size="10" name="yuhyoend" value="" onClick="jsPopCal('f','yuhyoend');" style="cursor:hand;" class="input_b"></td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����������</td>
	<td><input type="radio" name="fixyn" value="Y">��� <input type="radio" name="fixyn" value="N" checked>������</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>����� �߿���� ����</td>
	<td><input type="radio" name="importantnotice" value="Y">��� <input type="radio" name="importantnotice" value="N" checked>������</td>
</tr>
</form>
</table>
<br>
<input type="button" value=" ��� " onclick="SubmitForm()">
<br><br>
(��ũ ���)<br>
&lt;a href="javascript:GoParent('http://www.10x10.co.kr/event/eventmain.asp?eventid=2607')"&gt;�ҳ� �̺�Ʈ �ٷΰ���&lt;/a&gt;

<!-- #include virtual="/lib/db/dbclose.asp" -->