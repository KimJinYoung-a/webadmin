<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Page : /admin/eventmanage/event/pop_event_Board_xls.asp
' Description :  �̺�Ʈ �Խ��� ������ Excel �ɼǼ��� �˾�
' History : 2009.05.06 ������ ����
'###########################################################

dim eCode
eCode = Request("eC")	'�̺�Ʈ�ڵ�
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<title>�Խ��� ������ �ɼ� ����</title>
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
<!--
	function chkForm()
	{
		var frm = document.frmOption;

		if(frm.Sdate.value.length<10) {
			alert("�Խ��� ���� �������� �Է����ּ���.");
			frm.Sdate.focus();
			return false;
		}

		if(frm.Edate.value.length<10) {
			alert("�Խ��� ���� �������� �Է����ּ���.");
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
<form name="frmOption" method="get" onsubmit="return chkForm()" action="pop_event_Board_xls_Download.asp">
<tr height="23">
	<td colspan="2" bgcolor="#F3F3F5"><b>�̺�Ʈ �Խ��� ������ �ٿ�ε� �ɼ� ����</b></td>
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
<tr>
	<td bgcolor="#F8F8FA" align="center">ȸ�����</td>
	<td>
		<select name="limitLevel">
			<option value="all">��ü ������</option>
			<% If (now() >= #08/01/2018 00:00:00#) then %>
			<option value="white">white��� ����</option>
			<% else %>
			<option value="orange">Orange��� ����</option>
			<option value="yellow">Yellow���� ����</option>
			<% end if %>
		</select>
	</td>
</tr>
<tr height="23">
	<td colspan="2" bgcolor="#F5F5F8" align="center"><input type="submit" value="�ٿ�ε�"></td>
</tr>
</form>
</table>
</body>
</html>
