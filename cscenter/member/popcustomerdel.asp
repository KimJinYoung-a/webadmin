<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : cs���� ȸ��Ż��
' History : 2019.01.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPWithLog.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
dim userid, userseq, i
	userid = requestCheckvar(request("userid"),32)
	userseq = requestCheckvar(request("userseq"),32)

%>
<script type="text/javascript">

function DelonUser() {
    if (frm.complaintext.value==""){
        alert("Ż������� �Է��� �ּ���.");
        return;
    }

	if (confirm('�¶��� ���� Ż��ó�� �մϴ�.\nŻ���Ŀ��� ������������ ������ ���� �Ұ��� �մϴ�.\n�����Ͻðڽ��ϱ�?') == true) {
		frm.mode.value = "delonuser";
		frm.action = "/cscenter/member/domodifyuserinfo.asp";
		frm.submit();
	}
}

</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			�� �� Ż�� ó��
		</td>
		<td align="right">		
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<form name="frm" method="post" action="" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" value="">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="userseq" value="<%= userseq %>">
<table width="100%" border="0" align="center" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA">
<tr align="left">
    <td height="30" width="120" bgcolor="#DDDDFF">Ż����� :</td>
    <td bgcolor="#FFFFFF" >
        <textarea cols="100" rows="5" name="complaintext"></textarea>
    </td>
</tr>
<tr>
	<td align="center" colspan=2 bgcolor="#FFFFFF">
		<input type="button" class="button" value="Ż��ó��" onClick="DelonUser();">
		<input type="button" class="button" value=" â�ݱ� " onClick="self.close()">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
