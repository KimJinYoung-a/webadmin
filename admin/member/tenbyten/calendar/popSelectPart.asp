<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<link rel="stylesheet" href="/bct.css" type="text/css">
<script language="javascript">

function putItem() {
	var frm = document.frm;

	if (frm.department_id.value == "") {
		alert("부서를 선택하세요.");
		frm.department_id.focus();
		return;
	}

	var result = opener.addPartItem(frm.department_id.options[frm.department_id.selectedIndex].text, frm.department_id.value);
	if (result != "") {
		alert(result);
	}
}

</script>
<form name="frm" method="GET" action="">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
<tr height="10" valign="bottom" bgcolor="F4F4F4">
	<td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	<td background="/images/tbl_blue_round_02.gif"></td>
	<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
</tr>
<tr height="25" valign="bottom" bgcolor="F4F4F4">
	<td background="/images/tbl_blue_round_04.gif"></td>
	<td valign="top" bgcolor="F4F4F4"><b>부서 선택</b></td>
	<td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<table width="100%" border="0" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" width="80">부 서</td>
	<td>
		<%= drawSelectBoxDepartment("department_id", "") %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td bgcolor="F8F8F8" align="center" colspan="2" height="45">
		<input type="button" class="button" value=" 추 가 " onClick="putItem()">
		&nbsp;&nbsp;
		<input type="button" class="button" value=" 닫 기 " onClick="self.close()">
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/lib/db/dbclose.asp" -->
