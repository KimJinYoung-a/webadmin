<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� traffic analysis  ���� �Է� ������
' History : 2007.09.04 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/traffic/traffic_class.asp"-->

<script language="javascript">
function sudongsubmit()
{
document.frm.action = "traffic_analysis_sudong_submit.asp";
document.frm.submit();
}

function back()
{
history.back();
}

</script>

<!--ǥ ������-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong>�� �ٹ����� traffic analysis �����Է�</strong></font>
			</td>
			
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td></td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr  height="10" valign="top">
		<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
		<td background="/images/tbl_blue_round_06.gif"></td>
		<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
	</tr>
	</tr>
</table>
<!--ǥ ��峡-->

<!--���� ����-->
	<table width="100%" border="0" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA" align="center">
	<form action="" name="frm" method="get">
	<tr bgcolor=#FFFFFF>
		<td align="right" colspan=6><input type="button" value="�ٹ����� DB�� ����" onclick="sudongsubmit()"> <input type="button" value="����������" onclick="back()"></td>
	</tr>
		<tr bgcolor=#DDDDFF>
			<td align="center">��¥</td>
		   <td align="center">��������</td>
		   <td align="center">�湮�ڼ�</td>
		   <td align="center">�űԹ湮�ڼ�</td>
		   <td align="center">��湮�ڼ�</td>
		   <td align="center">�����湮�ڼ�</td>
		</tr>
		<tr bgcolor=#FFFFFF>
			<td align="center"><input type="text" maxsize="10" name="yyyymmdd"></td>
			<td align="center"><input type="text" maxsize="10" name="pageview"></td>
			<td align="center"><input type="text" maxsize="10" name="totalcount"></td>
			<td align="center"><input type="text" maxsize="10" name="newcount"></td>
			<td align="center"><input type="text" maxsize="10" name="recount"></td>
			<td align="center"><input type="text" maxsize="10" name="realcount"></td>
		</tr>
	</form>	
	</table>
<!-- ���� �� -->

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->