<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->

<% 
dim shopid : shopid = requestCheckvar(request("shopid"),32)
dim param1 : param1 = requestCheckvar(request("param1"),32)
dim param2 : param2 = requestCheckvar(request("param2"),32)
%>
<script language='javascript'>
function selThis(comp){
    var frm = comp.form;
    var shopid= frm.shopid.value;
    if (shopid.length<1){
        alert('���ó�� �����ϼ���.');
        return;
    }
    opener.popRetShopid(shopid,'<%=param1%>','<%=param2%>');
    window.close();
}

</script>
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="30" align="center" bgcolor="<%= adminColor("tabletop") %>">
	<form name="frm">
	    <td width="40">���ó</td>
    	<td><% drawSelectBoxOffShop "shopid",shopid %></td>
    <tr>
    <tr height="25" bgcolor="FFFFFF">
    	<td align="center" colspan="2">
    	    <input type="button" value="����" onClick="selThis(this)">
    	&nbsp;
    	    <input type="button" value="���" onClick="window.close();">
    	</td>
    <tr>
    </form>
</table>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->