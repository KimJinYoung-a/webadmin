<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �ؿ� ���� ��ǰ �������
' History : 2018.04.25 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/cscenter/lib/popheader.asp"-->

<%
dim orderserial, oUniPassNumber
	orderserial = requestCheckVar(request("orderserial"),16)

If orderserial <> "" And Not isnull(orderserial) Then
	oUniPassNumber = fnUniPassNumber(orderserial)
end if
%>
<script type="text/javascript">

document.title = "�ؿ� ���� ����";

function fnCustomNumberSubmit(){
	var frm =  document.frm;
	if(!frm.customNumber.value || frm.customNumber.value.length < 13){
		alert('13�ڸ��� �������������ȣ �� �Է� ���ּ���.');
		frm.customNumber.focus();
		return;
	}

	var str1 = frm.customNumber.value.substring(0,1);
	var str2 = frm.customNumber.value.substring(1,13);

	if((str1.indexOf("P") < 0) == true){
		alert('P�� �����ϴ� 13�ڸ� ��ȣ�� �Է� ���ּ���.');
		frm.customNumber.focus();
		return;
	}

	var regNumber = /^[0-9]*$/;
	if (!regNumber.test(str2)){
		alert('��ȣ�� ���ڸ� �Է����ּ���.');
		frm.customNumber.focus();
		return;
	}

	frm.mode.value = "editforeigndirectpurchase";
	frm.action = "/cscenter/ordermaster/order_info_edit_process.asp";
	frm.submit();
}

</script>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" onsubmit="return false;" action="">
<input type="hidden" name="mode" value="">
<input type="hidden" name="orderserial" value="<%=orderserial%>" />
<tr height="25" bgcolor="<%= adminColor("topbar") %>">
    <td colspan="2">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
    		<tr>
    			<td width="50%">
    				<img src="/images/icon_star.gif" align="absbottom">&nbsp;<b>�ؿ� ���� ����</b>
			    </td>
			    <td align="right">
			    	<input type="button" value="�����ϱ�" class="csbutton" onclick="fnCustomNumberSubmit();">
			    </td>
			</tr>
		</table>
    </td>
</tr>
<tr height="25" bgcolor="#FFFFFF">
    <td bgcolor="<%= adminColor("topbar") %>">������� ������ȣ</td>
    <td><input type="text" id="individualNum" name="customNumber" value="<%=oUniPassNumber%>" maxlength="14" size=14 /></td>
</tr>
</table>

<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
