<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ������ѵ��
' History : 2011.01.19 ������ ����
'			2017.09.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPwithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim userid
	userid = requestCheckVar(request("userid"),32)

If not(C_ADMIN_AUTH or C_PSMngPart) Then
	response.write "<script  type='text/javascript'>"
	response.write "	alert('������ �����ϴ�.');"
	response.write "</script>"
	dbget.close() : response.end
end if
%>

<script type="text/javascript">
	function jsSubmit(){
		if (jsChkBlank(document.frmPW.sPW.value)){
			alert("������ ��й�ȣ�� �Է��ϼ���.");
			document.frmPW.sPW.focus();
			return;
		}

		if (document.frmPW.sPW.value.replace(/\s/g, "").length < 6 || document.frmPW.sPW.value.replace(/\s/g, "").length > 16){
			alert("��й�ȣ�� ������� 6~16���Դϴ�.");
			document.frmPW.sPW.focus();
			return ;
		}

		if ((document.frmPW.sPW.value)!=(document.frmPW.sPW1.value)){
			alert("��й�ȣ�� ��ġ���� �ʽ��ϴ�.");
			document.frmPW.sPW1.focus();
			return;
		}

		if (!fnChkComplexPassword(frmPW.sPW.value)) {
			alert('���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
			frmPW.sPW.focus();
			return;
		}

		if(confirm("��к�ȣ�� �����Ͻðڽ��ϱ�?")){
			document.frmPW.submit();
		}

	}

	//�ε�� ��Ŀ��
	window.onload = function(){
		document.frmPW.sPW.focus();
	}
</script>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
		<br>�� �н����� ����� �ʱ�ȭ �Ǵ� ���
		<br>1. ������ ������ ���� �ΰ��, ��������� �����
		<br>2. ��Ⱓ �̻������ ���� ������ �����, ����� ������.
		<br>3. �н����带 Ʋ���� �����, ����� ���� �˴ϴ�.
	</td>
	<td align="right"></td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frmPW" method="post" action="/admin/member/tenbyten/procUseridChangedPw.asp">
	<input type="hidden" name="uid" value="<%=userid%>">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="28">
		<td width="100" bgcolor="#E6E6E6" align="center">�ٹ�����ID</td>
		<td bgcolor="#ffffff"><%=userid%></td>
	</tr>
	<tr>
		<td bgcolor="#E6E6E6"  align="center">��й�ȣ</td><td bgcolor="#ffffff"><input type="password" name="sPW" size="16">
			<div style="font-size:8pt;padding:1px;">���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)</div>
			</td>
	</tr>
	<tr>
		<td bgcolor="#E6E6E6"  align="center">��й�ȣ Ȯ��</td><td bgcolor="#ffffff"><input type="password" name="sPW1" size="16"></td>
	</tr>
</table>
<div style="width:100%;text-align:center;padding:10"><input type="button" class="button" value="Ȯ��" onClick="jsSubmit();"></div>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->