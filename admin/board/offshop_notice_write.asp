<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<script>
function SubmitForm()
{
        if (document.f.gubun.value == "") {
                alert("�����ֱ� ������ �����ϼ���.");
                return;
        }
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
<table border="1" cellpadding="0" cellspacing="0" bordercolordark="White" bordercolorlight="#808080" class="a">
<form method="post" name="f" action="offshop_notice_act.asp" onsubmit="return false" enctype="MULTIPART/FORM-DATA">
<input type="hidden" name="mode" value="write">
<input type="hidden" name="userid" value="<%=session("ssBctId")%>">
<input type="hidden" name="username" value="<%=session("ssBctCname")%>">
<tr>
	<td align="center">�����ֱ� ����</td>
	<td>
		<select name="gubun">
			<option value="">����</option>
			<option value="00">��ü</option>
			<option value="01">1F Shop</option>
			<option value="02">2F Zoom</option>
			<option value="03">3F College</option>
			<option value="04">�¶��λ����</option>
			<option value="50">����-��ü</option>
			<option value="51">����-����</option>
			<option value="52">����-������</option>
			<option value="53">����-������</option>
		</select>
	</td>
</tr>
<tr>
	<td align="center">����</td>
	<td><input type="text" name="title" size="30" value=""></td>
</tr>
<tr>
	<td align="center">����</td>
	<td><textarea name="contents" cols="50" rows="15"></textarea></td>
</tr>
<tr>
	<td align="center">����</td>
	<td><input type="file" name="file" size="30"></td>
</tr>
<tr><td colspan="2" align="right"><input type="button" value=" ��� " onclick="SubmitForm()">&nbsp;&nbsp;</td></tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->