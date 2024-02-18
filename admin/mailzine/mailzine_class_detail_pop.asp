<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ������
' History : 2018.04.27 �̻� ����(���Ϸ� ���� ���� ���Ϸ��� �߼� ���� ����. ���� �������� ����.)
'			2019.06.24 ������ ����(���ø� ��� �ű� �߰�)
'			2020.05.28 �ѿ�� ����(TMS ���Ϸ� �߰�)
'###########################################################
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/mailzinecls.asp"-->
<%

'// ============================================================================
dim classDate, mode

classDate = requestCheckVar(request("dt"),32)

if (classDate <> "") then
	mode = "modicls"
else
	mode = "inscls"
end if


'// ============================================================================
dim oClass
set oClass = new CMailzineList
oClass.FRectDate = classDate
oClass.frectmailergubun = "EMS"
oClass.MailzineClassOne()

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/jquery-1.7.2.min.js"></script>
<script language="JavaScript">
function jsSubmit(frm) {
	if (frm.classDate.value == '') {
		alert('�߼����ڸ� �Է��ϼ���.');
		return;
	}

	var day = new Date(frm.classDate.value);
	var today = new Date();

	/*
	if (frm.classDate.value <= today.yyyymmdd()) {
		alert('�߼����ڴ� �����ں��� �Է°����մϴ�.\n(���� �Ǵ� ������¥ �ԷºҰ�!)');
		return;
	}
	*/

	if (isInt(frm.itemid1.value, false) != true) {
		alert('Ŭ���� 01 �� ��ǰ�ڵ带 �Է��ϼ���.');
		return;
	}

	if (isInt(frm.salePer1.value, false) != true) {
		alert('Ŭ���� 01 �� �������� �Է��ϼ���.');
		return;
	}

	if (frm.classDesc1.value == '') {
		alert('Ŭ���� 01 �� ���¼����� �Է��ϼ���.');
		return;
	}

	if (frm.classSubDesc1.value == '') {
		alert('Ŭ���� 01 �� ���¼��꼳���� �Է��ϼ���.');
		return;
	}

	if (day.getDay() == 5) {
		// �ݿ����� Ŭ������ 3���̴�.
		if (isInt(frm.itemid2.value, false) != true) {
			alert('Ŭ���� 02 �� ��ǰ�ڵ带 �Է��ϼ���.');
			return;
		}

		if (isInt(frm.salePer2.value, false) != true) {
			alert('Ŭ���� 02 �� �������� �Է��ϼ���.');
			return;
		}

		if (frm.classDesc2.value == '') {
			alert('Ŭ���� 02 �� ���¼����� �Է��ϼ���.');
			return;
		}

		if (frm.classSubDesc2.value == '') {
			alert('Ŭ���� 02 �� ���¼��꼳���� �Է��ϼ���.');
			return;
		}

		if (isInt(frm.itemid3.value, false) != true) {
			alert('Ŭ���� 03 �� ��ǰ�ڵ带 �Է��ϼ���.');
			return;
		}

		if (isInt(frm.salePer3.value, false) != true) {
			alert('Ŭ���� 03 �� �������� �Է��ϼ���.');
			return;
		}

		if (frm.classDesc3.value == '') {
			alert('Ŭ���� 03 �� ���¼����� �Է��ϼ���.');
			return;
		}

		if (frm.classSubDesc3.value == '') {
			alert('Ŭ���� 03 �� ���¼��꼳���� �Է��ϼ���.');
			return;
		}
	} else {
		if ((frm.itemid2.value != '') || (frm.itemid3.value != '')) {
			if (confirm('�ݿ��ϸ� 3���� Ŭ���� ����� �����մϴ�.\n\n�Էµ� Ŭ���� 02/03 �� ������ ������� �ʽ��ϴ�.\n\n�����Ͻðڽ��ϱ�?') != true) {
				return;
			}
		}
	}

	if (confirm('�����Ͻðڽ��ϱ�?') == true) {
		frm.submit();
	}
}

Date.prototype.yyyymmdd = function() {
  var mm = this.getMonth() + 1; // getMonth() is zero-based
  var dd = this.getDate();

  return [this.getFullYear(),
          (mm>9 ? '' : '0') + mm,
          (dd>9 ? '' : '0') + dd
         ].join('-');
};

function isInt(value, allowBlank) {
	if ((value == '') && (allowBlank == true)) { return true; }
	return !isNaN(value) && parseInt(value, 10) == value;
}

function jsSetDisabledObj(obj, disabled) {
	obj.disabled = disabled;
	if (obj.type != 'textarea') {
		obj.style.background = disabled ? '#DDDDDD' : '#FFFFFF';
	}
}

function jsSetItemState() {
	var frm = document.frm;
	if (frm.classDate.value == '') { return; }
	var day = new Date(frm.classDate.value);

	if (day.getDay() == 5) {
		jsSetDisabledObj(frm.itemid2, false);
		jsSetDisabledObj(frm.salePer2, false);
		jsSetDisabledObj(frm.classDesc2, false);
		jsSetDisabledObj(frm.classSubDesc2, false);
		jsSetDisabledObj(frm.itemid3, false);
		jsSetDisabledObj(frm.salePer3, false);
		jsSetDisabledObj(frm.classDesc3, false);
		jsSetDisabledObj(frm.classSubDesc3, false);
	} else {
		jsSetDisabledObj(frm.itemid2, true);
		jsSetDisabledObj(frm.salePer2, true);
		jsSetDisabledObj(frm.classDesc2, true);
		jsSetDisabledObj(frm.classSubDesc2, true);
		jsSetDisabledObj(frm.itemid3, true);
		jsSetDisabledObj(frm.salePer3, true);
		jsSetDisabledObj(frm.classDesc3, true);
		jsSetDisabledObj(frm.classSubDesc3, true);
	}
}

$(document).ready(function(){
	jsSetItemState();
});

</script>
<table width="95%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
	<form name="frm" method="post" action="mailzine_process.asp">
	<input type="hidden" name="mode" value="<%= mode %>">
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" width="150">�߼�����</td>
		<td colspan="2">
			<input id="classDate" name="classDate" value="<%= oClass.FOneItem.FclassDate %>" class="text_ro" size="10" maxlength="10" readonly />
			<% if (mode = "inscls") then %>
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="classDate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script language="javascript">
			var classDate = new Calendar({
				inputField : "classDate", trigger    : "classDate_trigger",
				onSelect: function() {
					this.hide();
					jsSetItemState();
				}, bottomBar: true, dateFormat: "%Y-%m-%d", fdow: 0
			});
			</script>
			<% end if %>
		</td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" width="150" rowspan="4">Ŭ���� 01</td>
		<td align="center" width="100">��ǰ�ڵ�</td>
		<td align="left"><input type="text" class="text" name="itemid1" value="<%= oClass.FOneItem.Fitemid1 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">������</td>
		<td align="left"><input type="text" class="text" name="salePer1" value="<%= oClass.FOneItem.FsalePer1 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">���¼���</td>
		<td align="left"><input type="text" class="text" name="classDesc1" value="<%= oClass.FOneItem.FclassDesc1 %>" size="50"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">���¼��꼳��</td>
		<td align="left"><input type="text" class="text" name="classSubDesc1" value="<%= oClass.FOneItem.FclassSubDesc1 %>" size="50"></td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" width="150" rowspan="4">Ŭ���� 02</td>
		<td align="center" width="100">��ǰ�ڵ�</td>
		<td align="left"><input type="text" class="text" name="itemid2" value="<%= oClass.FOneItem.Fitemid2 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">������</td>
		<td align="left"><input type="text" class="text" name="salePer2" value="<%= oClass.FOneItem.FsalePer2 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">���¼���</td>
		<td align="left"><input type="text" class="text" name="classDesc2" value="<%= oClass.FOneItem.FclassDesc2 %>" size="50"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">���¼��꼳��</td>
		<td align="left"><input type="text" class="text" name="classSubDesc2" value="<%= oClass.FOneItem.FclassSubDesc2 %>" size="50"></td>
	</tr>

	<tr bgcolor="#FFFFFF" height="25">
		<td align="center" width="150" rowspan="4">Ŭ���� 03</td>
		<td align="center" width="100">��ǰ�ڵ�</td>
		<td align="left"><input type="text" class="text" name="itemid3" value="<%= oClass.FOneItem.Fitemid3 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">������</td>
		<td align="left"><input type="text" class="text" name="salePer3" value="<%= oClass.FOneItem.FsalePer3 %>" size="10"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">���¼���</td>
		<td align="left"><input type="text" class="text" name="classDesc3" value="<%= oClass.FOneItem.FclassDesc3 %>" size="50"></td>
	</tr>
	<tr bgcolor="#FFFFFF" height="25">
		<td align="center">���¼��꼳��</td>
		<td align="left"><input type="text" class="text" name="classSubDesc3" value="<%= oClass.FOneItem.FclassSubDesc3 %>" size="50"></td>
	</tr>
	</form>

	<tr bgcolor="#FFFFFF" height="50">
		<td align="center" colspan="3">
			<input type="button" class="button" value=" �� �� �� �� " onClick="jsSubmit(document.frm)">
		</td>
	</tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
