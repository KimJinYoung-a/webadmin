<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/util/datelib.asp"-->

<script language='javascript'>

function submitForm(upfrm){
	if (upfrm.useridarr.value == ""){
		alert("���̵� �Է����ּ���!");
		upfrm.useridarr.focus();
		return;
	}

	if(upfrm.couponvalue.value == ""){
		alert("�����ݾ� �Ǵ� �������� �Է����ּ���!");
		upfrm.couponvalue.focus();
		return;
	}

	if(upfrm.couponname.value == ""){
		alert("�������� �Է����ּ���!");
		upfrm.couponname.focus();
		return;
	}

	if(upfrm.minbuyprice.value == ""){
		alert("�ּұݾ��� �Է����ּ���!");
		upfrm.minbuyprice.focus();
		return;
	}

	if(upfrm.startdate.value == "" || upfrm.expiredate.value == ""){
		alert("���Ⱓ�� �Է����ּ���!");
		return;
	}

	if (confirm('���� ���ʽ� ������ �����Ͻðڽ��ϱ�?\n\n�ع���� ������ ����� �� �����Ƿ� ������ Ȯ���ϼ���.')){
		upfrm.submit();
	}
}

function EnableBox(comp){
	if (comp.checked){
		frmarr.targetitemlist.disabled = false;
		frmarr.couponmeaipprice.disabled = false;

		frmarr.targetitemlist.style.backgroundColor = "#FFFFFF";
		frmarr.couponmeaipprice.style.backgroundColor = "#FFFFFF";
	}else{
		frmarr.targetitemlist.disabled = true;
		frmarr.couponmeaipprice.disabled = true;

		frmarr.targetitemlist.style.backgroundColor = "#E6E6E6";
		frmarr.couponmeaipprice.style.backgroundColor = "#E6E6E6";
	}

}
</script>
<font color="#FF6699">*** �޸��� ����(ex : corpse2,icommang)</font>
<table width="760" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="#B2B2B2">
<form name="frmarr" method="post" action="lecCouponedit_Process.asp">
<input type="hidden" name="mode" value="">
<tr>
	<td bgcolor="#E6E6E6" width="130" align="center">���̵��߰�</td>
	<td width="200" align="right"><input type="text" name="useridarr" value="" size="80"></td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">����Ÿ��</td>
	<td bgcolor="#FFFFFF">
	<input type=text name=couponvalue maxlength=7 size=10>
	<input type=radio name=coupontype value="1" onclick="alert('% ���� �����Դϴ�.');">%����
	<input type=radio name=coupontype value="2" checked >������
	(�ݾ� �Ǵ� % ����)
	</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">������</td>
	<td bgcolor="#FFFFFF"><input type=text name=couponname maxlength="100" size=80>
	<br>(Ex. ���ȸ���� ���� Ư���� ���� 10%���� ����)</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�ּұ��űݾ�</td>
	<td bgcolor="#FFFFFF"><input type=text name=minbuyprice maxlength=7 size=10>�� �̻� ���� ������ ��밡��(����)</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">��ȿ�Ⱓ</td>
	<td bgcolor="#FFFFFF"><input type=text name=startdate value="<%= left(now(),10) %> 00:00:00" maxlength=19 size=19>~<input type=text name=expiredate maxlength=19 size=19>(<%= left(now(),10) %> 00:00:00 ~ <%= left(now(),10) %> 23:59:59)</td>
</tr>
<tr>
	<td bgcolor="#E6E6E6" align="center">�߱��� ID </td>
	<td bgcolor="#FFFFFF"><%= session("ssBctId") %></td>
</tr>
<tr>
	<td colspan="2" align=center bgcolor="#FFFFFF"><input type=button value="����" onClick="submitForm(this.form);"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->