<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%

dim idx, bankdate
idx = request("idx")
bankdate = request("bankdate")

%>

<script language="javascript">

function jsSubmitMatch(frm) {
	if (frm.finishstr.value == "") {
		alert("������ �Է��ϼ���.");
		frm.finishstr.focus();
		return;
	}

	if (getByteLength(frm.finishstr.value) > 32) {
		alert("������ �ʹ� �� �Է��� �� �����ϴ�.\n(�ѱ۱��� 16�ڱ��� ����)");
		frm.finishstr.focus();
		return;
	}

	if (frm.ipkumCause.value == "") {
		alert("�Աݻ����� �����ϼ���.");
		frm.ipkumCause.focus();
		return;
	}

	if ((frm.ipkumCause.value == "�����Է�") && (frm.ipkumCauseText.value == "")) {
		alert("�Աݻ����� �Է��ϼ���.");
		frm.ipkumCauseText.focus();
		return;
	}

	if (frm.ipkumCause.value == "�����Է�") {
		if (getByteLength(frm.ipkumCauseText.value) > 32) {
			alert("�Աݻ����� �ʹ� �� �Է��� �� �����ϴ�.\n(�ѱ۱��� 16�ڱ��� ����)");
			frm.ipkumCauseText.focus();
			return;
		}
	}

	if (confirm("��Ī�Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function Change_ipkumCause(comp) {
    if (comp.value=="�����Է�") {
        document.all.span_ipkumCauseText.style.display = "inline";
    }else{
        document.all.span_ipkumCauseText.style.display = "none";
    }
}

function getByteLength(str) {
	var ret;

	ret = 0;
	for (var i = 0; i <= str.length - 1; i++) {
		var ch = str.charAt(i);
		if (escape(ch).length > 4) {
			ret = ret + 2;
		} else {
			ret = ret + 1;
		}
	}

    return ret;
}

</script>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="post" action="pop_matchorderlist_Process.asp">
	<input type="hidden" name="mode" value="matchByHand">
	<input type="hidden" name="ipkumidx" value="<%= idx %>">
	<input type="hidden" name="bankdate" value="<%= bankdate %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
		<td colspan="2">�����Ī �����Է�(�ֹ���ȣ ��)</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
		<td width="80">����</td>
    	<td align="left">
			<input type="text" class="text" name="finishstr" size="25" value="">
		</td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" height="30">
		<td width="80">�Աݻ���</td>
    	<td align="left">
			<select class="select" name="ipkumCause" onChange="Change_ipkumCause(this);">
				<option value=""></option>
				<option value="�߰� ��ۺ�">�߰� ��ۺ�</option>
				<option value="�߰� ��ǰ���">�߰� ��ǰ���</option>
				<option value="�ֹ�����">�ֹ�����</option>
				<option value="���� ����">���� ����</option>
				<option value="�����Է�">�����Է�</option>
			</select>
			<span name="span_ipkumCauseText" id="span_ipkumCauseText" style='display:none'>
			<input type="text" class="text" name="ipkumCauseText" size="15" value="">
		</td>
    </tr>
	</form>
</table>

<br>

<div align="center">
<input type="button" class="button" value="��Ī�ϱ�" onClick="jsSubmitMatch(frm)">
</div>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
