<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� ����
' History : 2012.12.11 �̻� ����
'			2013.04.23 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<!-- #include virtual="/lib/classes/offshopclass/offmanualmeachulcls.asp"-->

<%

dim ErrStr

%>

<script language='javascript'>

String.prototype.trim = function() {
    return this.replace(/(^\s*)|(\s*$)/gi, "");
}

function checkClick() {
	var frm = document.frm;
	var orgdata = frm.orgdata.value;
	var oneline, onelineItems, blankLineCount;

	orgdata = orgdata.split("\n");
	blankLineCount = 0;

	for (var i = 0; i < orgdata.length; i++) {
		oneline = orgdata[i];

		if (oneline.trim() == "") {
			blankLineCount = blankLineCount + 1;
			continue;
		}

		onelineItems = oneline.split("\t");

		if (onelineItems.length != 14) {
			alert("������ ������ 14���� �÷��� �Ǿ�� �մϴ�.\n\n" + oneline);
			return false;
		}

		// ����
		if (onelineItems[1] != "90") {
			alert("90 ��ǰ�� ��ϰ����մϴ�.");
			return false;
		}

		if (onelineItems[1] != "90") {
			alert("90 ��ǰ�� ��ϰ����մϴ�. [" + onelineItems[1] + "]");
			return false;
		}

		if ((onelineItems[4] == "") || (onelineItems[5] == "") || (onelineItems[6] == "")) {
			alert("ī�װ��� �����Ǿ����ϴ�.");
			return false;
		}

		if (onelineItems[7].replace(",", "")*0 != 0) {
			alert("�ݾ��� ���ڰ� �ƴմϴ�. [" + onelineItems[7] + "]");
			return false;
		}

		if ((onelineItems[12] != "����") && (onelineItems[12] != "�鼼")) {
			alert("�߸��� ���������Դϴ�.");
			return false;
		}
	}

	if ((orgdata.length - blankLineCount) > 100) {
		alert("�ѹ��� 100���� ��ǰ���� ��ϰ����մϴ�.");
		return false;
	}

	alert("OK : ��ǰ�� " + (orgdata.length - blankLineCount));

	return true;
}

function uploadClick() {
	var frm = document.frm;

	if (checkClick() != true) {
		return;
	}

	if (confirm('�ڷḦ ���ε��մϴ�.\n\n�����Ͻðڽ��ϱ�?')) {
		frm.mode.value="uploaddata";
		frm.submit();
	}
}

function saveClick() {
	var frm = document.frm;
	var checkeditemexist = false;
	var dataarr = "";

	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			break;
		}

		if (v.checked == true) {
			checkeditemexist = true;
			break;
		}
	}

	if (checkeditemexist == false) {
		alert("������ ������ �����ϴ�.");
		return;
	}

	if (confirm("������ �Ͻðڽ��ϱ�?") == true) {
		dataarr = "-1";
		for (var i = 0; ; i++) {
			var v = document.getElementById("chk_" + i);
			if (v == undefined) {
				break;
			}

			if (v.checked == true) {
				dataarr = dataarr + "," + v.value
			}
		}

		frm.orgdata.value = dataarr;

		frm.mode.value="regtemporder";
		frm.submit();
	}
}

function delClick() {
	var frm = document.frm;
	var checkeditemexist = false;
	var dataarr = "";

	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			break;
		}

		if (v.checked == true) {
			checkeditemexist = true;
			break;
		}
	}

	if (checkeditemexist == false) {
		alert("������ ����� �����ϴ�.");
		return;
	}

	if (confirm("���� �Ͻðڽ��ϱ�?") == true) {
		dataarr = "-1";
		for (var i = 0; ; i++) {
			var v = document.getElementById("chk_" + i);
			if (v == undefined) {
				break;
			}

			if (v.checked == true) {
				dataarr = dataarr + "," + v.value
			}
		}

		frm.orgdata.value = dataarr;

		frm.mode.value="deltemporder";
		frm.submit();
	}
}

function CheckAll(chk) {
	for (var i = 0; ; i++) {
		var v = document.getElementById("chk_" + i);
		if (v == undefined) {
			return;
		}

		if (v.disabled != true) {
			v.checked = chk.checked;
		}
	}
}

function clearData() {
	var frm = document.frm;
	frm.orgdata.value = "";
}

</script>

<table border=0 cellspacing=0 cellpadding=0 class="a">
<form name="frm" method="post" action="<%=uploadImgUrl%>/linkweb/offshop/item/item_add_multi_off.asp" onSubmit="return false;">
<input type="hidden" name="mode" value="">
<tr>
	<td>
		<font color="red">������ �и�</font><br>
		�귣��ID, ��ǰ����, ��ǰ��, �ɼǸ�, ī�װ�(��), ī�װ�(��), ī�װ�(��), �ǸŰ�, ���԰�, ������ް�, �������, ���͸��Ա���, ��������, ������ڵ�<br>
		TOMS001&lt;TAB>90&lt;TAB>NAVY CANVAS CLASSICS TINY&lt;TAB>120/5&lt;TAB>070&lt;TAB>030&lt;TAB>060&lt;TAB>25,000&lt;TAB>0&lt;TAB>0&lt;TAB>�����&lt;TAB>Ư��&lt;TAB>����&lt;TAB>TSAA90000000296<br>
		<font color="red">�ɼ��� ������ ��簪���� ������ ������ ����� �ȵ˴ϴ�.</font>

		<% if (ErrStr <> "") then %>
			<br><br><b><font color="red"><%= Replace(ErrStr, "\n", "<br>") %></font></b><br><br>
		<% end if %>
	</td>
	<td align="right" valign="bottom">
	</td>
</tr>
<tr>
	<td colspan=2>
	<textarea name="orgdata" cols=120 rows=5></textarea>
	</td>
</tr>
<tr>
	<td>
	<input type= button class="button" value="Clear" onClick="clearData();">
	</td>
	<td>
		<input type= button class="button" value=" ü ũ " onclick="checkClick()">
		<input type= button class="button" value="���ε�" onclick="uploadClick()">
	</td>
</tr>
</form>
</table>

<p>

</form>
</table>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
