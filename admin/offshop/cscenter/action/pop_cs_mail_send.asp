<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �������� ������
' Hieditor : 2011.03.08 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/popheader_cs_off.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/email/maillib.asp" -->
<!-- #include virtual="/admin/offshop/cscenter/cscenter_Function_off.asp"-->
<%
dim email, masteridx
	email 	= requestCheckVar(request("email"),128)
	masteridx 	= requestCheckVar(request("masteridx"),10)
%>

<script type='text/javascript'>

function SendCSMail(mailform){

	if (mailform.mailto.value.length<1){
		alert('�����ּҸ� �Է��ϼ���.');
		return;
	}
	if (mailform.title.value.length<1){
		alert('�������� �Է��ϼ���.');
		return;
	}
	if (mailform.contents.value.length<1){
		alert('���ϳ��븦 �Է��ϼ���.');
		return;
	}

	var ret= confirm('���� �Ͻðڽ��ϱ�?');
	
	if(ret){
		mailform.submit();
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
    	<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>���Ϲ߼�</b>
    </td>
    <td background="/images/tbl_blue_round_05.gif"></td>
</tr>
</table>
<!-- ǥ ��ܹ� ��-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<form name="mailform" method="post" action="pop_cs_mail_send_process.asp">
<input type="hidden" name="masteridx" value="<%= masteridx %>">
<tr>
	<td width="60" bgcolor="<%= adminColor("tabletop") %>">�����ּ�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="mailto" class="text" value="<%= email %>" size="30"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">��������</td>
	<td bgcolor="#FFFFFF"><input type="text" name="title" class="text" value="" size="80"></td>
</tr>
<tr>
	<td bgcolor="<%= adminColor("tabletop") %>">���ϳ���</td>
	<td bgcolor="#FFFFFF"><textarea name="contents" class="textarea" value="" cols="80" rows="19"></textarea></td>
</tr>    
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
<tr valign="bottom" height="25">
    <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
    <td valign="bottom" align="center">
        <input type="button" class="button" value="���Ϲ߼�" onclick="SendCSMail(mailform);">
    </td>
    <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
</tr>
<tr valign="top" height="10">
    <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
    <td background="/images/tbl_blue_round_08.gif"></td>
    <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
</tr>
</form>
</table>
<!-- ǥ �ϴܹ� ��-->

<!-- #include virtual="/admin/offshop/cscenter/poptail_cs_off.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->