<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim prect, rect

prect = requestCheckVar(request("prect"), 32)
rect = requestCheckVar(request("rect"), 32)

%>

<script language='javascript'>

function jsSubmit(frm) {
	if (frm.UserAddCNT.value == "") {
		alert('����ġ�� �Է��ϼ���.');
		return;
    }

	if (frm.UserAddCNT.value*0 != 0) {
		alert('����ġ�� ���ڸ� �����մϴ�.');
		return;
    }

	var ret = confirm("���� �Ͻðڽ��ϱ�?");
	if(ret){
		frm.submit();
	}
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�����˻��� ���</b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method=post action="manageRelatedKeyword_process.asp">
	<input type="hidden" name="mode" value="modievtUserAddCNT">
	<input type="hidden" name="prect" value="<%= prect %>">
	<input type="hidden" name="rect" value="<%= rect %>">
    <tr>
    	<td width="100" bgcolor="<%= adminColor("tabletop") %>">���˻���</td>
    	<td bgcolor="#FFFFFF">
    		<%= prect %>
    	</td>
    </tr>
    <tr>
    	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�����˻���</td>
    	<td bgcolor="#FFFFFF">
    		<%= rect %>
    	</td>
    </tr>
    <tr>
    	<td width="100" bgcolor="<%= adminColor("tabletop") %>">����ġ</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" class="text" name="UserAddCNT" value="" size="10">
    	</td>
    </tr>
    </form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <input type="button" class="button" value="����" onclick="jsSubmit(frm);">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->