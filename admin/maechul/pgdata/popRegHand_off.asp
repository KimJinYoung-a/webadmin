<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� �Ǹ� ��� ����
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->
<%

%>
<script src="/cscenter/js/jquery-1.7.1.min.js"></script>
<script language="javascript">

function jsSubmit(){
	var frm = document.frm;

	if ($("input[name=orgpgkey]").val() == '') {
		alert('PG��KEY �� �Է��ϼ���.');
		return;
	}

	if ($("input[name=gubun]:checked").val() == undefined) {
		alert('������ �����ϼ���.');
		return;
	}

	if ($("input[name=gubun]:checked").val() == 'cancel') {
		if ($("input[name=canceldate]").val() == '') {
			alert('����Ͻø� �Է��ϼ���.');
			return;
		}

		var fromDate = new Date('<%= Left(DateAdd("m", -1, Now()), 7) + "-01" %>');
		var toDate = new Date('<%= Left(DateAdd("m", 1, Now()), 7) + "-01" %>');
		var cancelDate = new Date($("input[name=canceldate]").val());

		if (isNaN(cancelDate)) {
			alert('�߸��� ��������Դϴ�.');
			return;
		} else if ((cancelDate < fromDate) || (cancelDate >= toDate)) {
			alert('�߸��� ��������Դϴ�.(' + formatDate(cancelDate) + ')');
			return;
		}

		/*
		if ($("input[name=ipkumdate]").val() == '') {
			alert('�Աݿ������� �Է��ϼ���.');
			return;
		}
		*/
	}

	frm.submit();
}

function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
}

</script>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<b>���ⵥ��Ÿ ���(OFF)</b>
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<form name="frm" method="post" action="pgdata_process.asp" style="margin: 0px;">
<input type="hidden" name="mode" value="addhand" />
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">��PG��KEY:</td>
	<td align="left">
		<input type="text" class="text" name="orgpgkey" size="32">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">����:</td>
	<td align="left">
		<input type="radio" name="gubun" value="cancel"> ī������
		&nbsp;
		<input type="radio" name="gubun" value="dup"> �ݾ� 0�� ���ΰ�
		&nbsp;
		<input type="radio" name="gubun" value="del"> ������ ����
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">����(���)�Ͻ�:</td>
	<td align="left">
		<input type="text" class="text" name="canceldate" size="32"> * ��: 2019-02-19 17:46:52
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�Աݿ�����:</td>
	<td align="left">
		<input type="text" class="text" name="ipkumdate" size="10">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td width="100" bgcolor="<%= adminColor("tabletop") %>">�ֹ���ȣ:</td>
	<td align="left">
		<input type="text" class="text" name="orderserial" size="10">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF">
	<td align="center" colspan="2" height="35">
	    <input type="button" class="button" value="���" onClick="jsSubmit();">
	    <input type="button" class="button" value="���" onClick="self.close();">
	</td>
</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
