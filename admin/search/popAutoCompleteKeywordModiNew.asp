<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

dim prect, rect
dim mode

rect = requestCheckVar(request("rect"), 32)

if (rect = "") then
	mode = "add"
else
	mode = "modievtUserAddCNT"
end if

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

function jsSubmitAdd(frm) {
	if (frm.rect.value == "") {
		alert('Ű���带 �Է��ϼ���.');
		return;
    }

	var ret = confirm("��� �Ͻðڽ��ϱ�?");
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
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>Ű���� ���</b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" method=post action="manageAutoCompleteKeyword_process.asp">
	<input type="hidden" name="mode" value="<%= mode %>">
	<% if (mode = "add") then %>
    <tr>
    	<td width="100" bgcolor="<%= adminColor("tabletop") %>">Ű����</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" class="text" name="rect" value="" size="10">
    	</td>
    </tr>
	<% else %>
	<input type="hidden" name="rect" value="<%= rect %>">
    <tr>
    	<td width="100" bgcolor="<%= adminColor("tabletop") %>">Ű����</td>
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
	<% end if %>
    </form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
			<% if (mode = "add") then %>
            <input type="button" class="button" value="�߰�" onclick="jsSubmitAdd(frm);">
			<% else %>
			<input type="button" class="button" value="����" onclick="jsSubmit(frm);">
			<% end if %>
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
