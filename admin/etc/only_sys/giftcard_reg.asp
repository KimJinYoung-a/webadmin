<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/etc/only_sys/check_auth.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<script>
function jsGiftCardReg(){
	if(frm1.iid.value == ""){
		alert("�ٹ����� Gift ī�带 �����ϼ���.");
		frm1.iid.focus();
		return;
	}
	if(frm1.opt.value == ""){
		alert("�ݾױ��� �����ϼ���.");
		frm1.opt.focus();
		return;
	}
	if(frm1.mmstitle.value == ""){
		alert("MMS ������ �Է��ϼ���.");
		frm1.mmstitle.focus();
		return;
	}
	if(frm1.mmsmessage.value == ""){
		alert("MMS �޼����� �Է��ϼ���.");
		frm1.mmsmessage.focus();
		return;
	}
	if(frm1.userid.value == ""){
		alert("���̵� �Է��ϼ���.");
		frm1.userid.focus();
		return;
	}
	frm1.submit();
}
</script>

<form name="frm1" action="giftcard_reg_proc.asp" method="post" style="margin:0px;">
<table class="a">
<tr>
	<td>
		<table cellpadding="0" cellspacing="0" border="0" class="a">
		<tr>
			<td height="50">
				����Ʈī���ȣ : <select name="iid"><option value="">-����-</option><option value="101" selected>[101]�ٹ����� Gift ī��</option></select>
				&nbsp;&nbsp;&nbsp;
				<select name="opt">
					<option value="">-�ݾױǼ���-</option>
					<option value="0001">1������</option>
					<option value="0002">2������</option>
					<option value="0003">3������</option>
					<option value="0004">5������</option>
					<option value="0005">8������</option>
					<option value="0006">10������</option>
					<option value="0007">15������</option>
					<option value="0008">20������</option>
					<option value="0009">30������</option>
				</select>
			</td>
		</tr>
		<tr>
			<td height="50">
				MMS ���� : <font color="red">�� ' " ���� �ʼ�.</font><br>
				<input type="text" name="mmstitle" id="mmstitle" value="" size="70"><br><br>
			</td>
		</tr>
		<tr>
			<td height="50">
				MMS �޼��� : <font color="red">�� ' " ���� �ʼ�.</font><br>
				<textarea name="mmsmessage" id="mmsmessage" rows="4" cols="100">[�ٹ�����] ����� �����ϸ����� ���� ��ǰ�ı� �̺�Ʈ�� ��÷�Ǽ̽��ϴ�.
��÷�ǽ� �в��� �ٹ����� ����Ʈ ī�� 1�������� �帳�ϴ�.
����Ʈ ī��� �����ٹ����ٿ��� Ȯ�� �����մϴ�.</textarea><br><br>
			</td>
		</tr>
		<tr>
			<td height="50">
				�߱� ���̵� : <br>
				<textarea name="userid" id="userid" rows="10" cols="100"></textarea><br><br>
				<input type="button" class="button" value="�� ��" style="width:100px;height:60px;" onClick="jsGiftCardReg()">
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->