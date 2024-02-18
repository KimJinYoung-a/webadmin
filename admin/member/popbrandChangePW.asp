<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��й�ȣ ����
' History : 2014.02.03 ������ ����
'			2021.07.16 �ѿ�� ����(2���н����� ����)
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/checkAllowIPwithLog.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%
	Dim brandid, sType
	brandid = requestCheckVar(request("bid"),32)
	sType		= requestCheckVar(request("sT"),1)
	'������ ���� Ȯ��
	if not (C_ADMIN_AUTH or C_SYSTEM_Part or C_CSUser or C_MD or C_OFF_part or C_logics_Part) then
			Call Alert_close ("�����ڸ� ���氡���մϴ�.������ Ȯ�����ּ���")
	end if
%>

<script type="text/javascript">
	function jsSubmit(){
		if (jsChkBlank(document.frmPW.sPW.value)){
			alert("������ ��й�ȣ�� �Է��ϼ���.");
			document.frmPW.sPW.focus();
			return;
		}

		if (document.frmPW.sPW.value.replace(/\s/g, "").length < 8 || document.frmPW.sPW.value.replace(/\s/g, "").length > 16){
			alert("��й�ȣ�� ������� 8~16���Դϴ�.");
			document.frmPW.sPW.focus();
			return ;
		}

		if ((document.frmPW.sPW.value)!=(document.frmPW.sPW1.value)){
			alert("��й�ȣ�� ��ġ���� �ʽ��ϴ�.");
			document.frmPW.sPW1.focus();
			return;
		}

		//if (jsChkBlank(document.frmPW.sPWS1.value)){
		//	alert("������ 2�� ��й�ȣ�� �Է��ϼ���.");
		//	document.frmPW.sPWS1.focus();
		//	return;
		//}

		//if (document.frmPW.sPWS1.value.replace(/\s/g, "").length < 8 || document.frmPW.sPWS1.value.replace(/\s/g, "").length > 16){
		//	alert("��й�ȣ�� ������� 8~16���Դϴ�.");
		//	document.frmPW.sPWS1.focus();
		//	return ;
		//}

		//if ((document.frmPW.sPWS1.value)!=(document.frmPW.sPWS2.value)){
		//	alert("��й�ȣ�� ��ġ���� �ʽ��ϴ�.");
		//	document.frmPW.sPWS1.focus();
		//	return;
		//}

		if (!fnChkComplexPassword(frmPW.sPW.value)) {
			alert('�н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
			frmPW.sPW.focus();
			return;
		}
		//if (!fnChkComplexPassword(frmPW.sPWS1.value)) {
		//	alert('2�� ���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)');
		//	frmPW.sPWS1.focus();
		//	return;
		//}

		if(confirm("��к�ȣ�� �����Ͻðڽ��ϱ�?")){
			document.frmPW.submit();
		}

	}

	//�ε�� ��Ŀ��
	window.onload = function(){
		document.frmPW.sPW.focus();
	}
</script>

<form name="frmPW" method="post" action="/admin/member/procbrandChangePW.asp" style="margin:0px;">
<input type="hidden" name="bid" value="<%=brandid%>">
<input type="hidden" name="sT" value="<%=sType%>">
<table width="100%" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="28">
		<td width="100" bgcolor="#E6E6E6" align="center">�귣��ID</td>
		<td bgcolor="#ffffff"><%=brandid%></td>
	</tr>
	<tr>
		<td bgcolor="#E6E6E6"  align="center">��й�ȣ</td><td bgcolor="#ffffff"><input type="password" name="sPW" size="16">
			<div style="font-size:8pt;padding:1px;">���ο� �н������ ����/����/Ư������ �� �ΰ��� �̻��� �������� �Է��ϼ���. �ּұ��� 10��(2����) , 8��(3����)</div>
			</td>
	</tr>
	<tr>
		<td bgcolor="#E6E6E6"  align="center">��й�ȣ Ȯ��</td><td bgcolor="#ffffff"><input type="password" name="sPW1" size="16"></td>
	</tr>
	<!--<tr>
		<td bgcolor="#E6E6E6"  align="center">2�� ��й�ȣ</td><td bgcolor="#ffffff"><input type="password" name="sPWS1" size="16"> 
			</td>
	</tr>-->
	<!--<tr>
		<td bgcolor="#E6E6E6"  align="center">2�� ��й�ȣ Ȯ��</td><td bgcolor="#ffffff"><input type="password" name="sPWS2" size="16"></td>
	</tr>-->
</table>
<div style="width:100%;text-align:center;padding:10"><input type="button" class="button" value="Ȯ��" onClick="jsSubmit();"></div>
</form>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->