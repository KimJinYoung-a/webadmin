<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/academy/lib/academy_function.asp"-->
<!-- #include virtual="/academy/lib/classes/fingers_ordercls.asp"-->
<%
Dim idx, oorderdetail
Dim entryname, entryhp, mode
	idx = RequestCheckvar(request("idx"),10)
	entryname 	= RequestCheckvar(request("entryname"),32)
	entryhp 	= RequestCheckvar(request("entryhp"),16)
	mode		= RequestCheckvar(request("mode"),16)
If mode = "U" Then
	Dim sqlStr

	sqlStr = sqlStr & " update [db_academy].[dbo].tbl_academy_order_detail set "
	sqlStr = sqlStr & " entryname = '"&entryname&"', entryhp = '"&entryhp&"' "
	sqlStr = sqlStr & " where detailidx = '"&idx&"' "
	dbACADEMYget.Execute sqlStr
	
	response.write "<script>alert('�����Ǿ����ϴ�.');</script>"
	response.write "<script>opener.location.reload();</script>"
	response.write "<script>window.close();</script>"
	response.End
End If

set oorderdetail = new CLectureFingerOrder
	oorderdetail.FRectidx = idx
	oorderdetail.GetFingerOrderDetailOne
%>
<script language="javascript">
function updateInfo(frm){

	if (frm.entryname.value== ""){
		alert('������ �Է��ϼ���.');
		frm.entryname.focus();
		return;
    }

	if (frm.entryhp.value.length<12){
		alert('����ó�� ��Ȯ�� �Է��ϼ���.');
		frm.entryhp.focus();
		return;
    }

	var ret= confirm("������ �����ϰڽ��ϱ�?");
	if(ret){
		frm.submit();
	}
}
window.resizeTo('580','500');
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
        	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>��������</b>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <form name="frm" action="pop_namephone_update.asp">
    <input type="hidden" name="idx" value="<%=idx%>">
    <input type="hidden" name="mode" value="U">
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">����</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="entryname" value="<%= oorderdetail.FItemList(0).Fentryname %>" size="13" maxlength="16">
    	</td>
    </tr>
    <tr>
    	<td width="60" bgcolor="<%= adminColor("tabletop") %>">����ó</td>
    	<td bgcolor="#FFFFFF">
    		<input type="text" name="entryhp" value="<%= oorderdetail.FItemList(0).Fentryhp %>" size="15" maxlength="16">
    	</td>
    </form>
</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center">
            <input type="button" class="button" value="��������" onclick="updateInfo(frm);">
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
<!-- #include virtual="/lib/db/dbAcademyclose.asp" -->