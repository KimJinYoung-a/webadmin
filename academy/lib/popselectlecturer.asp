<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<%
dim lecuserid, parentcomp
lecuserid = RequestCheckvar(request("lecuserid"),32)
parentcomp = RequestCheckvar(request("parentcomp"),16)
%>

<script language='javascript'>
self.resizeTo(400,200);

function selectLec(frm){
	if (frm.lecuserid.value.length<1){
		alert('강사아이디를 선택하세요.');
		frm.lecuserid.focus();
		return;
	}

	opener.<%= parentcomp %>.value = frm.lecuserid.value;
	self.close();
}
</script>
<table width="340" height="100" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
<form name="frm">
<tr>
	<td align="center">
		<% drawSelectBoxLecturer "lecuserid", lecuserid %><input type="button" value="선택" onclick="selectLec(frm)">
	</td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->