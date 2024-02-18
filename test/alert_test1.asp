<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTadmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<script language='javascript'>
function ss(){
    if (confirm('ok?')){
		document.frm.target = "xLink";
		document.frm.action = "<%=apiURL%>/outmall/testAlert.asp"
		document.frm.submit();
    }
}
</script>

<table>
<form name="frm" method="post" action="">
<tr>
    <td>
        <input type="text" name = "ttt">
    </td>
    <td>
        <input type="button" value="click" onclick="ss();">
    </td>
</tr>
</form>
</table>
<iframe name="xLink" id="xLink" frameborder="0" width="100%" height="100"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
