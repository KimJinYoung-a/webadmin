<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Page : /admin/eventmanage/eventWinner/event_EntryList.asp
' Description :  �̺�Ʈ ��÷�� ���� �ۼ� ������
' History : 2007.09.27 ������
'####################################################
%>


<%


dim replyName,replyMail,mailTitle,mailContents

	replyName = request("rpName")
	replyMail = request("rpMail")
	mailTitle = request("mlTitle")
	mailContents = request("mlCont")
	if mailContents<>"" then mailContents=Replace(mailContents, vbcrlf,"<br>")


dim fso,contFile,MailPath,MailForm

	MailPath = server.mappath("/lib/email/email_event.htm")

	set fso = Server.Createobject("Scripting.filesystemObject")
	set contFile = fso.Opentextfile(MailPath)

	MailForm = contFile.readAll

	contFile.close

	MailForm= replace(MailForm,"$$MAILCONTENTS$$",mailContents)

	set fso = nothing

%>

<script>

</script>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>�ٹ����� ���� ���� �̸�����</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<META content="MSHTML 5.50.4937.800" name=GENERATOR>
<Style>
.a {font:9pt/135% "����";color:#000000}
</style>
</HEAD>
<BODY style="FONT-SIZE: 9pt; COLOR: #000000; FONT-FAMILY: ����; BACKGROUND-COLOR: #ffffff" bgColor=#ffffff leftMargin=0 background="" topMargin=0 marginheight="0" marginwidth="0">

<table align="center" width="600" border="0" cellpadding="0" cellspacing="0" class="a" bgcolor="#F4F4F4">
	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td width="10" background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<!-- ���� ��� �⺻���� ���� -->
        	<table width="100%" border="0" cellpadding="0" cellspacing="1" class="a" bgcolor="#BABABA">
				<tr bgcolor="#FFFFFF">
					<td align="center" width="120">������ �̸� </td><td align="left" width="420"><%= replyName %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td align="center" width="120">������ ���� �ּ�</td><td align="left"><%= replyMail %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td align="center" width="120">���� ����</td><td align="left"><%= mailTitle %></td>
				</tr>
			</table>
			<!-- // ���� ��� �⺻���� �� -->
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top" style="padding : 0 0 10 0">
        <td width="10" background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	<!-- ������ ���� -->
			<%= MailForm %>
			<!-- ������ �� -->
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>

</BODY>
</HTML>



