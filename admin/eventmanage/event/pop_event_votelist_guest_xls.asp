<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Comment_xls.asp
' Description :  �̺�Ʈ �ڸ�Ʈ ������ Excel �ɼǼ��� �˾�
' History : 2007.10.12 ������ ����
'###########################################################

dim eCode
eCode = Request("eC")	'�̺�Ʈ�ڵ�

rsget.open "SELECT COUNT(*) FROM [db_event].[dbo].[tbl_event_subscript] WHERE evt_code = '" & eCode & "'",dbget,1
IF rsget(0) = 0 Then
	Response.Write "<script>alert('�����Ͱ� �����ϴ�.');window.close();</script>"
	dbget.close()
	Response.End
End IF
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title>������ �ɼ� ����</title>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
<!--
	function chkForm()
	{
		var frm = document.frmOption;

		if(frm.Sdate.value.length<10) {
			alert("���� �������� �Է����ּ���.");
			frm.Sdate.focus();
			return false;
		}

		if(frm.Edate.value.length<10) {
			alert("���� �������� �Է����ּ���.");
			frm.Edate.focus();
			return false;
		}

		if(confirm("�����Ͻ� �ɼ����� Excel������ �ٿ�ε��Ͻðڽ��ϱ�?")) {
			return true;
		}
		else {
			return false;
		}
	}
//-->
</script>
</head>
<body style="margin:0px 0px 0px 0px;">
<table width="400" cellpadding="2" cellspacing="2" border="0" class="a">
<form name="frmOption" method="get" onsubmit="return chkForm()" action="pop_event_vote_xls_Download_guest.asp">
<tr height="23">
	<td colspan="2" bgcolor="#F3F3F5"><b>�̺�Ʈ ������ �ٿ�ε� �ɼ� ����</b></td>
</tr>
<tr>
	<td width="100" bgcolor="#F8F8FA" align="center">�̺�Ʈ �ڵ�</td>
	<td>
		<%=eCode%>
		<input type="hidden" name="eC" value="<%=eCode%>">
	</td>
</tr>
<tr>
	<td bgcolor="#F8F8FA" align="center">�����Ⱓ</td>
	<td>
		<input type="text" name="Sdate" size="10" maxlength="10">
		~
		<input type="text" name="Edate" size="10" maxlength="10">
		<br>�� ��) 2007-10-12 ~ 2007-10-15
	</td>
</tr>
<tr height="23">
	<td colspan="2" bgcolor="#F5F5F8" align="center"><input type="submit" value="�ٿ�ε�"></td>
</tr>
</form>
</table>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->