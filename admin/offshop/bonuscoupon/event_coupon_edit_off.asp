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

	if ((upfrm.coupontype[0].checked == true) && (upfrm.couponvalue.value*1 > 15)) {
		// ������
		alert('15% �� �Ѵ� ���������� ������ �� �����ϴ�.');
		upfrm.couponvalue.focus();
		return;
	}

	if ((upfrm.coupontype[1].checked == true) && (upfrm.couponvalue.value*1 > upfrm.minbuyprice.value*0.2)) {
		// ������
		alert('�������ξ��� �ּұ��űݾ��� 20% �� ���� �� �����ϴ�.');
		upfrm.couponvalue.focus();
		return;
	}

//���� 2006-05-09 �ּ�ó��
//	if (upfrm.targetitemusing.checked){
//		if (!IsDigit(upfrm.targetitemlist.value)) {
//			alert('��ǰ��ȣ�� ���ڸ� �����մϴ�.');
//			upfrm.targetitemlist.focus();
//			return;
//		}

//		if ((upfrm.couponmeaipprice.value!='')&&(!IsDigit(upfrm.couponmeaipprice.value))) {
//			alert('���԰��� ���ڸ� �����մϴ�.');
//			upfrm.couponmeaipprice.focus();
//			return;
//		}
//	}

	if (confirm('������ �����Ͻðڽ��ϱ�?')){
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

		frmarr.targetitemlist.style.backgroundColor = "<%= adminColor("tabletop") %>";
		frmarr.couponmeaipprice.style.backgroundColor = "<%= adminColor("tabletop") %>";
	}

}
</script>

<table width="760" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="#B2B2B2">
<form name="frmarr" method="post" action="eventcouponedit_Process_off.asp">
<input type="hidden" name="mode" value="">
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" width="130" align="center">���̵��߰�</td>
	<td align="left">
		<input type="text" class="text" name="useridarr" value="" size="40">
		<font color="#FF6699">*** �޸��� ����(ex : corpse2,icommang)</font>
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">����Ÿ��</td>
	<td bgcolor="#FFFFFF">
	<input type="text" class="text" name=couponvalue maxlength=7 size=10>
	<input type=radio name=coupontype value="1" onclick="alert('% ���� �����Դϴ�.');">%����
	<input type=radio name=coupontype value="2" checked >������
	(�ݾ� �Ǵ� % ����)
	</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">������</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name=couponname maxlength="100" size=80>
	<br>(2004�� 1�� ���� 5������ 10���� �̻� ���Ű����� �帮�� ��ǰ���Դϴ�.)</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�ּұ��űݾ�</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name=minbuyprice maxlength=7 size=10>�� �̻� ���Ž� ��밡��(����)</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">��ȿ�Ⱓ</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name=startdate value="<%= left(now(),10) %> 00:00:00" maxlength=19 size=20> ~ <input type="text" class="text" name=expiredate maxlength=19 size=20>(<%= left(now(),10) %> 00:00:00 ~ <%= left(now(),10) %> 23:59:59)</td>
</tr>
<tr height=30>
	<td bgcolor="<%= adminColor("tabletop") %>" align="center">�߱��� ID </td>
	<td bgcolor="#FFFFFF"><%= session("ssBctId") %></td>
</tr>
<tr height=30>
	<td colspan="2" align=center bgcolor="#FFFFFF"><input type=button class=button value="����" onClick="submitForm(this.form);"></td>
</tr>
</form>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->