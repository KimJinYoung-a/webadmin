<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description : ��ǰ���
' History : ������ ����
'			2017.04.12 �ѿ�� ����(���Ȱ���ó��)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<%
dim designer,react
	designer = requestCheckVar(request("designer"),32)
	react = requestCheckVar(request("react"),10)

response.write "<script type='text/javascript'>location.href='/admin/offshop/popoffitemreg.asp?makerid=" + designer + "';</script>"
dbget.close()	:	response.End

%>
<script type='text/javascript'>

function refreshParent(){
	opener.frm.submit();
}

function AddOffItem(frm){
	if ((frm.itemgubun[0].checked==false)&&(frm.itemgubun[1].checked==false)){
		alert('��ǰ������ �����ϼ���.');
		return;
	}

	if (frm.designer.value.length<1){
		alert('�귣�带 �����ϼ���.');
		return;
	}

	if (frm.itemname.value.length<1){
		alert('�����۸��� �ϼ���.');
		frm.itemname.focus();
		return;
	}

	if (!IsDigit(frm.sellcash.value)){
		alert('�ǸŰ��� ���ڸ� �����մϴ�.');
		frm.sellcash.focus();
		return;
	}

	if (!IsDigit(frm.suplycash.value)){
		alert('��ü ���԰��� ���ڸ� �����մϴ�.');
		frm.suplycash.focus();
		return;
	}

	if (!IsDigit(frm.shopbuyprice.value)){
		alert('�� ���ް��� ���ڸ� �����մϴ�.');
		frm.shopbuyprice.focus();
		return;
	}

	if (((frm.suplycash.value!=0)||(frm.shopbuyprice.value!=0))){
		if (!confirm('!! �⺻ ��� ������ �ٸ� ��쿡�� ���԰� ���ް��� �Է� �ϼž� �մϴ�. \n\n��� �Ͻðڽ��ϱ�?')){
			return;
		}
	}

	var ret = confirm('�߰��Ͻðڽ��ϱ�?');

	if (ret) {
		frm.submit();
	}
}
</script>

<div align="center">
<br>
<table width="400" cellspacing="1" class="a" bgcolor=#3d3d3d>
<form name="frmadd" method="post" action="shopitem_process.asp">
<input type="hidden" name="mode" value="offitemreg">
<tr bgcolor="#FFFFFF">
	<td colspan=2>
	<table border=0 cellspacing=0 cellpadding=0 class="a" >
	<tr>
		<td width=110>�ۿ����� �����ǰ </td>
		<td>:�¶��� ��ǰ�� ������ ���� �� ���.</td>
	</tr>
	<tr>
		<td>���̺�Ʈ��ǰ </td>
		<td>:��ü �⺻ ���޸����� ���޸����� �ٸ����.<br><b>(���ް� ���� �Է�)</b></td>
	</tr>
	<tr>
		<td>�ۼҸ�ǰ </td>
		<td>:��Ÿ �Ҹ�ǰ.</td>
	</tr>
	<tr>
		<td>�۰����������ǰ </td>
		<td>:���������� ���� �Ǹ��ϴ»�ǰ.</td>
	</tr>
	</table>
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">����</td>
	<td bgcolor="#FFFFFF">
	<input type="radio" name="itemgubun" value="90">������ �����ǰ(90)<br>
	<input type="radio" name="itemgubun" value="70">�Ҹ�ǰ(70)<br>
	<!--
	<input type="radio" name="itemgubun" value="80" disabled >�̺�Ʈ��ǰ(80) : ������<br>
	<input type="radio" name="itemgubun" value="95" disabled >�����������ǰ(95) : ������
	-->
	</td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">�귣�� ID</td>
	<td bgcolor="#FFFFFF"><% drawSelectBoxDesignerwithName "designer",designer  %></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">��ǰ��</td>
	<td bgcolor="#FFFFFF"><input type="text" name="itemname" value="" maxlength="32"></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">�ǸŰ���</td>
	<td bgcolor="#FFFFFF"><input type="text" name="sellcash" value="" maxlength="9"></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">���԰�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="suplycash" value="0" size=6 maxlength="9"><br><b>(0�ϰ�� ��� ������ ���� �ڵ������˴ϴ�.)</b></td>
</tr>
<tr>
	<td width="100" bgcolor="#DDDDFF">�ް��ް�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="shopbuyprice" value="0" size=6 maxlength="9"><br><b>(0�ϰ�� ��� ������ ���� �ڵ������˴ϴ�.)</b></td>
</tr>
<tr>
	<td colspan="2" align="center" bgcolor="#FFFFFF"><input type="button" value="�߰�" onclick="AddOffItem(frmadd)"></td>
</tr>
</form>
</table>
</div>

<% if react="true" then %>
<!-- <script type='text/javascript'>refreshParent();</script> -->
<% end if %>

<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->