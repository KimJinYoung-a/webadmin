<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<script language="javascript">

function jsSubmitIns(frm) {
	if (frm.arrItemid.value == "") {
		alert("������ �Է��ϼ���.");
		return;
	}

	if (confirm("�����Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="itemImageText_Process.asp">
	<input type="hidden" name="mode" value="ins">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td>��û ��ǰ�ڵ�</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
    	<td>
			<textarea class="textarea" name="arrItemid" cols="15" rows="12"></textarea>
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value="�Է��ϱ�" onClick="jsSubmitIns(frm)">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
