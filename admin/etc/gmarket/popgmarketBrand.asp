<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim gMakername, gBrandname
%>
<script language="javascript">
function fnSaveForm() {
	var frm = document.frmAct;
    frm.target = "xLink2";
    frm.cmdparam.value = "AddMakerBrand";
    frm.action = "<%=apiURL%>/outmall/gmarket/actgmarketReq.asp"
    frm.submit();
}
</script>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmAct" target="xLink">
<input type="hidden" name="cmdparam">
<tr bgcolor="#FFFFFF">
	<td>브랜드ID</td>
	<td>
		<input type="text" name="makerid" size="50" maxlength=10  value="toms">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>제조사명</td>
	<td>
		<input type="text" name="gMakername" size="50" maxlength=10  value="<%=gMakername%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td>브랜드명</td>
	<td>
		<input type="text" name="gBrandname" size="50" maxlength=10  value="<%=gBrandname%>">
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td colspan="2" align="center"><input type="button" class="button" onclick="fnSaveForm();" value="검색"></td>
</tr>
</form>
</table>
<iframe name="xLink2" id="xLink2" frameborder="0" width="100%" height="300"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
